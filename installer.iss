; 編譯：ISCC /DAppVersion=<版本> installer.iss（版本號取自 Cargo.toml，見 README）
; 前置：先 cargo build --release
#ifndef AppVersion
  #error 請用 /DAppVersion=x.y.z 傳入版本號
#endif

[Setup]
; AppId 固定，之後升級才會被視為同一個程式
AppId={{6F1D2B7A-3C54-4E0B-9A8E-C11B0A5D7E42}
AppName=ClipboardApp
AppVersion={#AppVersion}
AppPublisher=YuCMochi
AppPublisherURL=https://github.com/YuCMochi/Clipboard
DefaultDirName={autopf}\ClipboardApp
DefaultGroupName=ClipboardApp
; 每使用者安裝：不用系統管理員權限，程式本來也只寫 HKCU
PrivilegesRequired=lowest
; 偵測到程式執行中會提示關閉，exe 才不會被鎖住
AppMutex=ClipboardApp_SingleInstance_Mutex
SetupIconFile=icon\app\256x256.ico
UninstallDisplayIcon={app}\ClipboardApp.exe
OutputDir=dist
OutputBaseFilename=ClipboardApp-Setup
Compression=lzma2
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64compatible
DisableProgramGroupPage=yes

[Tasks]
Name: "autostart"; Description: "開機自動啟動 (Start with Windows)"

[Files]
Source: "target\release\ClipboardApp.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\ClipboardApp"; Filename: "{app}\ClipboardApp.exe"

[Registry]
; 格式與程式內「開機自動啟動」開關相同，兩邊互通
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "ClipboardApp"; ValueData: """{app}\ClipboardApp.exe"" --autostart"; Tasks: autostart

[Run]
Filename: "{app}\ClipboardApp.exe"; Description: "立即啟動 ClipboardApp"; Flags: nowait postinstall skipifsilent

[Code]
// Run 值可能是程式內的開關寫的而非安裝程式，所以解除安裝時一律清掉
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usPostUninstall then
  begin
    RegDeleteValue(HKCU, 'Software\Microsoft\Windows\CurrentVersion\Run', 'ClipboardApp');
    RegDeleteKeyIncludingSubkeys(HKCU, 'Software\ClipboardApp');
  end;
end;
