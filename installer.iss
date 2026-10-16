; 編譯:ISCC /DAppVersion=<版本> installer.iss(版本號取自 Cargo.toml,見 README)
; 前置：先 cargo build --release
#ifndef AppVersion
  #error 請用 /DAppVersion=x.y.z 傳入版本號
#endif

[Setup]
; AppId 固定，之後升級才會被視為同一個程式
AppId={{B3A7C0E1-5D2F-4A68-8E19-7F4C2D9A6B05}
AppName=Pathie
AppVersion={#AppVersion}
AppPublisher=YuCMochi
AppPublisherURL=https://github.com/YuCMochi/Pathie
DefaultDirName={autopf}\Pathie
DefaultGroupName=Pathie
; 每使用者安裝：不用系統管理員權限，程式本來也只寫 HKCU
PrivilegesRequired=lowest
; 偵測到程式執行中會提示關閉，exe 才不會被鎖住
AppMutex=Pathie_SingleInstance_Mutex
SetupIconFile=icon\app\256x256.ico
UninstallDisplayIcon={app}\Pathie.exe
OutputDir=dist
OutputBaseFilename=Pathie-Setup
Compression=lzma2
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64compatible
DisableProgramGroupPage=yes

[Tasks]
Name: "autostart"; Description: "開機自動啟動 (Start with Windows)"

[Files]
Source: "target\release\Pathie.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\Pathie"; Filename: "{app}\Pathie.exe"

[Registry]
; 格式與程式內「開機自動啟動」開關相同，兩邊互通
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "Pathie"; ValueData: """{app}\Pathie.exe"" --autostart"; Tasks: autostart

[Run]
Filename: "{app}\Pathie.exe"; Description: "立即啟動 Pathie"; Flags: nowait postinstall skipifsilent

[Code]
// Run 值可能是程式內的開關寫的而非安裝程式，所以解除安裝時一律清掉
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usPostUninstall then
  begin
    RegDeleteValue(HKCU, 'Software\Microsoft\Windows\CurrentVersion\Run', 'Pathie');
    RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Pathie');
  end;
end;
