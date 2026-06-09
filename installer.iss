#define MyAppName "CikkChecker"
#define MyAppVersion "2.0.0"
#define MyAppPublisher "Sanyi7511"
#define MyAppURL "https://github.com/Sanyi7511/CikkChecker-releases"
#define MyAppExeName "CikkChecker.exe"

[Setup]
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}

; Telepítési mappa: C:\Users\...\AppData\Local\CikkChecker
DefaultDirName={localappdata}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes

; Kimenet
OutputDir=Output
OutputBaseFilename=CikkCheckerSetup
SetupIconFile=assets\icon.ico

Compression=lzma2/ultra64
SolidCompression=yes

; Windows 10+ 64-bit
MinVersion=10.0
ArchitecturesInstallIn64BitMode=x64compatible
ArchitecturesAllowed=x64compatible

; Nem kell rendszergazda
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog

WizardStyle=modern
WizardResizable=yes

; Eltávolítás a Programok közül
UninstallDisplayName={#MyAppName}
UninstallDisplayIcon={app}\{#MyAppExeName}
CreateUninstallRegKey=yes

[Languages]
Name: "hungarian"; MessagesFile: "compiler:Languages\Hungarian.isl"

[Tasks]
Name: "desktopicon"; Description: "Asztali parancsikon létrehozása"; GroupDescription: "További beállítások:"; Flags: checked

[Files]
; Fő alkalmazás exe (PyInstaller által generált)
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

; Rust backend bináris
Source: "checker_core.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
; Start menü
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{#MyAppName} eltávolítása"; Filename: "{uninstallexe}"

; Asztali ikon (csak ha a felhasználó kérte)
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; Opcionálisan elindítja telepítés végén
Filename: "{app}\{#MyAppExeName}"; Description: "CikkChecker indítása"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Törli az adatfájlokat eltávolításkor
Type: filesandordirs; Name: "{app}"
