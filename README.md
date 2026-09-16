# ClipboardApp

ClipboardApp 是一款 Windows 桌面工具，會在系統工具列常駐，監控你的剪貼簿內容，當你複製有效的檔案或資料夾路徑時，會自動幫你開啟該路徑。可在系統工具列左鍵點擊圖示快速暫停/恢復監控，或右鍵退出程式。

以 Rust 實作，直接呼叫 Win32 API，不需要 .NET runtime，產出單一原生 exe。

## 安裝與使用

1. 點擊 release ，下載並解壓縮 ClipboardApp.zip
2. 執行 `ClipboardApp.exe`，程式會自動縮到系統工具列。
3. 複製任何有效的檔案或資料夾路徑，程式會自動開啟。

支援的路徑形式：一般絕對路徑、加引號的路徑、正斜線路徑、`file:///` URL、含環境變數的路徑（如 `%TEMP%`）。

## 建置 (Build)

需要 [Rust](https://rustup.rs/) 與 MSVC build tools：

```
cargo build --release
```

產物在 `target/release/ClipboardApp.exe`，單一檔案即可執行。

跑測試：

```
cargo test
```

## 注意事項 (Notes)

- 請勿重複執行多個實例，程式已自動防呆。
- 程式是免安裝的單一 exe。
- 基於安全考量，只有**絕對路徑**才會觸發開啟；裸檔名（如 `gpedit.msc`）與相對路徑會被忽略，避免靠工作目錄湊巧命中系統指令。
