// ClipboardApp：常駐系統匣，偵測剪貼簿中的有效路徑並自動開啟。
//
// cfg_attr 是必要的：直接寫 windows_subsystem 會連 `cargo test` 的測試執行檔
// 一起變成 windows subsystem，測試輸出就看不到了。
#![cfg_attr(not(test), windows_subsystem = "windows")]

use std::cell::Cell;
use std::path::{Path, PathBuf};

use windows::core::{w, PCWSTR, PWSTR};
use windows::Win32::Foundation::{
    GetLastError, ERROR_ALREADY_EXISTS, HMODULE, HWND, LPARAM, LRESULT, POINT, WPARAM,
};
use windows::Win32::System::DataExchange::{
    AddClipboardFormatListener, RemoveClipboardFormatListener,
};
use windows::Win32::System::Environment::ExpandEnvironmentStringsW;
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::System::Threading::CreateMutexW;
use windows::Win32::UI::Shell::{
    PathCreateFromUrlW, ShellExecuteW, Shell_NotifyIconW, NIF_ICON, NIF_MESSAGE, NIF_TIP, NIM_ADD,
    NIM_DELETE, NIM_MODIFY, NOTIFYICONDATAW,
};
use windows::Win32::UI::WindowsAndMessaging::*;
use winreg::enums::{HKEY_CURRENT_USER, KEY_SET_VALUE};
use winreg::RegKey;

// 沿用 C# 版的登錄檔位置與值名，兩版之間的開機自動啟動設定可以互通
const RUN_KEY: &str = r"Software\Microsoft\Windows\CurrentVersion\Run";
const RUN_VALUE: &str = "ClipboardApp";

const WM_TRAY: u32 = WM_APP + 1;
const TRAY_UID: u32 = 1;

const ID_STARTUP: usize = 1;
const ID_TOGGLE: usize = 2;
const ID_EXIT: usize = 3;

// 對應 app.rc 的資源 ID
const IDI_ON: u16 = 2;
const IDI_OFF: u16 = 3;

thread_local! {
    static MONITORING: Cell<bool> = const { Cell::new(true) };
    static ICON_ON: Cell<HICON> = Cell::new(HICON::default());
    static ICON_OFF: Cell<HICON> = Cell::new(HICON::default());
    // explorer.exe 重啟時廣播的訊息；收到要重新加入托盤圖示，否則圖示消失但程式還在跑
    static WM_TASKBAR_CREATED: Cell<u32> = const { Cell::new(0) };
}

fn main() {
    unsafe {
        // 單一實例。沿用 C# 版的 mutex 名稱，所以新舊版本之間也會互斥。
        // handle 不釋放，活到行程結束。
        let _mutex = CreateMutexW(None, true, w!("ClipboardApp_SingleInstance_Mutex"));
        if GetLastError() == ERROR_ALREADY_EXISTS {
            return;
        }

        let hinst: HMODULE = GetModuleHandleW(None).expect("GetModuleHandleW");
        let class = w!("ClipboardAppWnd");
        let wc = WNDCLASSW {
            lpfnWndProc: Some(wndproc),
            hInstance: hinst.into(),
            lpszClassName: class,
            ..Default::default()
        };
        RegisterClassW(&wc);

        // 一般視窗但從不 ShowWindow，對齊 C# 隱藏 Form 的做法。
        // 不用 HWND_MESSAGE：message-only window 搭剪貼簿監聽有已知的雷。
        let hwnd = CreateWindowExW(
            WINDOW_EX_STYLE(0),
            class,
            w!("ClipboardApp"),
            WS_OVERLAPPED,
            0,
            0,
            0,
            0,
            None,
            None,
            hinst,
            None,
        )
        .expect("CreateWindowExW");

        let cx = GetSystemMetrics(SM_CXSMICON);
        let cy = GetSystemMetrics(SM_CYSMICON);
        ICON_ON.set(load_icon(hinst, IDI_ON, cx, cy));
        ICON_OFF.set(load_icon(hinst, IDI_OFF, cx, cy));

        WM_TASKBAR_CREATED.set(RegisterWindowMessageW(w!("TaskbarCreated")));
        let _ = AddClipboardFormatListener(hwnd);
        tray_add(hwnd);

        let mut msg = MSG::default();
        // > 0：GetMessageW 出錯時回 -1，用 as_bool() 會變成無窮迴圈
        while GetMessageW(&mut msg, HWND::default(), 0, 0).0 > 0 {
            let _ = TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }
}

unsafe extern "system" fn wndproc(hwnd: HWND, msg: u32, wp: WPARAM, lp: LPARAM) -> LRESULT {
    match msg {
        WM_CLIPBOARDUPDATE => {
            if MONITORING.get() {
                handle_clipboard();
            }
            LRESULT(0)
        }
        WM_TRAY => {
            match lp.0 as u32 {
                WM_LBUTTONUP => toggle_monitoring(hwnd),
                WM_RBUTTONUP => show_menu(hwnd),
                _ => {}
            }
            LRESULT(0)
        }
        WM_COMMAND => {
            match (wp.0 & 0xFFFF) as usize {
                ID_STARTUP => set_startup(!is_startup_enabled()),
                ID_TOGGLE => toggle_monitoring(hwnd),
                ID_EXIT => {
                    let _ = DestroyWindow(hwnd);
                }
                _ => {}
            }
            LRESULT(0)
        }
        WM_DESTROY => {
            let _ = Shell_NotifyIconW(NIM_DELETE, &tray_data(hwnd));
            let _ = RemoveClipboardFormatListener(hwnd);
            PostQuitMessage(0);
            LRESULT(0)
        }
        m if m != 0 && m == WM_TASKBAR_CREATED.get() => {
            tray_add(hwnd);
            LRESULT(0)
        }
        _ => DefWindowProcW(hwnd, msg, wp, lp),
    }
}

// --- 剪貼簿 ---------------------------------------------------------------

fn handle_clipboard() {
    // clipboard-win 內建重試：WM_CLIPBOARDUPDATE 當下來源程式往往還佔著剪貼簿，
    // 直接 OpenClipboard 會拿到 ACCESS_DENIED
    let Ok(text) = clipboard_win::get_clipboard_string() else {
        return;
    };
    let Some(path) = normalize(&text) else {
        return;
    };
    if is_valid(&path) {
        open_path(&path);
    }
}

fn normalize(raw: &str) -> Option<PathBuf> {
    let t = raw.trim().trim_matches('"');
    if t.is_empty() {
        return None;
    }

    // get(..7) 而非 &t[..7]：剪貼簿是任意文字，切在多位元組字元中間會 panic
    let t = match t.get(..7) {
        Some(p) if p.eq_ignore_ascii_case("file://") => {
            url_to_path(t).unwrap_or_else(|| t.to_owned())
        }
        _ => t.to_owned(),
    };

    let t = expand_env(&t.replace('/', "\\"));
    let t = t.trim();
    (!t.is_empty()).then(|| PathBuf::from(t))
}

fn is_valid(p: &Path) -> bool {
    // 只接受絕對路徑，避免裸檔名（如 gpedit.msc）靠工作目錄湊巧命中系統指令。
    // Rust 的 is_absolute 比 C# 的 Path.IsPathRooted 更嚴格，連 `\Windows\...`
    // 和 `C:foo` 這種「有 root 但非絕對」的寫法也一併擋掉。
    p.is_absolute() && p.exists()
}

fn open_path(p: &Path) {
    let wide: Vec<u16> = p.to_string_lossy().encode_utf16().chain(Some(0)).collect();
    unsafe {
        ShellExecuteW(
            HWND::default(),
            PCWSTR::null(),
            PCWSTR(wide.as_ptr()),
            PCWSTR::null(),
            PCWSTR::null(),
            SW_SHOWNORMAL,
        );
    }
}

/// `new Uri(x).LocalPath` 的對應 API：percent-decoding 與 UNC 都交給 shlwapi，
/// 省掉一個 url crate。
fn url_to_path(url: &str) -> Option<String> {
    let src: Vec<u16> = url.encode_utf16().chain(Some(0)).collect();
    let mut buf = vec![0u16; 4096];
    let mut len = buf.len() as u32;
    unsafe {
        PathCreateFromUrlW(PCWSTR(src.as_ptr()), PWSTR(buf.as_mut_ptr()), &mut len, 0).ok()?;
    }
    Some(String::from_utf16_lossy(&buf[..len as usize]))
}

fn expand_env(s: &str) -> String {
    let src: Vec<u16> = s.encode_utf16().chain(Some(0)).collect();
    unsafe {
        let n = ExpandEnvironmentStringsW(PCWSTR(src.as_ptr()), None);
        if n == 0 {
            return s.to_owned();
        }
        let mut buf = vec![0u16; n as usize];
        let written = ExpandEnvironmentStringsW(PCWSTR(src.as_ptr()), Some(&mut buf));
        if written == 0 {
            return s.to_owned();
        }
        // 回傳值含結尾的 null
        String::from_utf16_lossy(&buf[..(written as usize).saturating_sub(1)])
    }
}

// --- 系統匣 ---------------------------------------------------------------

fn tray_data(hwnd: HWND) -> NOTIFYICONDATAW {
    NOTIFYICONDATAW {
        cbSize: std::mem::size_of::<NOTIFYICONDATAW>() as u32,
        hWnd: hwnd,
        uID: TRAY_UID,
        ..Default::default()
    }
}

unsafe fn tray_add(hwnd: HWND) {
    let mut nid = tray_data(hwnd);
    nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    nid.uCallbackMessage = WM_TRAY;
    nid.hIcon = if MONITORING.get() {
        ICON_ON.get()
    } else {
        ICON_OFF.get()
    };
    let tip: Vec<u16> = "Clipboard 路徑自動開啟".encode_utf16().collect();
    nid.szTip[..tip.len()].copy_from_slice(&tip);
    let _ = Shell_NotifyIconW(NIM_ADD, &nid);
}

// ponytail: on/off 只進 32x32，16x16 讓 Windows 縮。兩個尺寸是各自獨立的 .ico
// 檔，不合併成單一多尺寸 .ico 就沒辦法讓 LoadImageW 自動選（C# 版同樣是拿
// 32x32 交給 WinForms 縮，行為對等）。要更銳利就把兩個尺寸合併成一個 .ico。
unsafe fn load_icon(hinst: HMODULE, id: u16, cx: i32, cy: i32) -> HICON {
    LoadImageW(
        hinst,
        PCWSTR(id as usize as *const u16),
        IMAGE_ICON,
        cx,
        cy,
        LR_DEFAULTCOLOR,
    )
    .map(|h| HICON(h.0))
    .unwrap_or_else(|_| LoadIconW(None, IDI_APPLICATION).unwrap_or_default())
}

unsafe fn toggle_monitoring(hwnd: HWND) {
    let on = !MONITORING.get();
    MONITORING.set(on);

    let mut nid = tray_data(hwnd);
    nid.uFlags = NIF_ICON;
    nid.hIcon = if on { ICON_ON.get() } else { ICON_OFF.get() };
    let _ = Shell_NotifyIconW(NIM_MODIFY, &nid);
}

unsafe fn show_menu(hwnd: HWND) {
    let Ok(menu) = CreatePopupMenu() else { return };

    let startup = if is_startup_enabled() {
        MF_CHECKED
    } else {
        MF_UNCHECKED
    };
    let _ = AppendMenuW(menu, MF_STRING | startup, ID_STARTUP, w!("開機自動啟動"));
    let toggle = if MONITORING.get() {
        w!("暫停監控")
    } else {
        w!("恢復監控")
    };
    let _ = AppendMenuW(menu, MF_STRING, ID_TOGGLE, toggle);
    let _ = AppendMenuW(menu, MF_STRING, ID_EXIT, w!("退出"));

    let mut pt = POINT::default();
    let _ = GetCursorPos(&mut pt);
    // 先搶前景、結束後補一則訊息，少了這兩步選單點到外面不會消失（經典 Win32 陷阱）
    let _ = SetForegroundWindow(hwnd);
    let _ = TrackPopupMenu(menu, TPM_RIGHTBUTTON, pt.x, pt.y, 0, hwnd, None);
    let _ = PostMessageW(hwnd, WM_NULL, WPARAM(0), LPARAM(0));
    let _ = DestroyMenu(menu);
}

// --- 開機自動啟動 ---------------------------------------------------------

fn exe_path() -> String {
    std::env::current_exe()
        .map(|p| p.to_string_lossy().into_owned())
        .unwrap_or_default()
}

fn is_startup_enabled() -> bool {
    let Ok(key) = RegKey::predef(HKEY_CURRENT_USER).open_subkey(RUN_KEY) else {
        return false;
    };
    let Ok(val) = key.get_value::<String, _>(RUN_VALUE) else {
        return false;
    };
    let exe = exe_path();
    !exe.is_empty() && val.trim_matches('"').eq_ignore_ascii_case(&exe)
}

fn set_startup(enable: bool) {
    let Ok(key) =
        RegKey::predef(HKEY_CURRENT_USER).open_subkey_with_flags(RUN_KEY, KEY_SET_VALUE)
    else {
        return;
    };
    if enable {
        let _ = key.set_value(RUN_VALUE, &format!("\"{}\"", exe_path()));
    } else {
        let _ = key.delete_value(RUN_VALUE);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn rejects_non_paths() {
        // commit b4b5317 要擋的就是這些：裸檔名、相對路徑、有 root 但非絕對
        let cases = [
            "gpedit.msc",
            "..\\Windows",
            "hello world",
            "",
            "   ",
            "\\Windows\\System32",
            "C:notrooted",
        ];
        for s in cases {
            let rejected = normalize(s).map_or(true, |p| !is_valid(&p));
            assert!(rejected, "應拒絕: {s:?}");
        }
    }

    #[test]
    fn multibyte_input_does_not_panic() {
        // 剪貼簿是任意文字；前綴比對若用 &t[..7] 會切在中文字中間 panic
        for s in ["複製一段中文", "檔", "路徑自動開啟", "日本語のテキスト"] {
            let _ = normalize(s);
        }
    }

    #[test]
    fn accepts_real_path_in_every_form() {
        let dir = std::env::temp_dir().join("clipboard_rs_test 空格");
        std::fs::create_dir_all(&dir).unwrap();
        let s = dir.to_string_lossy().into_owned();
        let slashed = s.replace('\\', "/");

        let forms = [
            s.clone(),
            format!("  {s}  "),
            format!("\"{s}\""),
            slashed.clone(),
            format!("file:///{}", slashed.replace(' ', "%20")),
        ];
        for f in &forms {
            let p = normalize(f).unwrap_or_else(|| panic!("normalize 回傳 None: {f:?}"));
            assert!(is_valid(&p), "應接受: {f:?} -> {p:?}");
        }
        let _ = std::fs::remove_dir_all(&dir);
    }

    #[test]
    fn expands_environment_variables() {
        let p = normalize("%TEMP%").expect("%TEMP% 應正規化成功");
        assert!(!p.to_string_lossy().contains('%'), "未展開: {p:?}");
        assert!(is_valid(&p), "%TEMP% 應展開成有效路徑: {p:?}");
    }
}
