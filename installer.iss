#define MyAppName "CikkChecker"
#define MyAppVersion "0.0.0"
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
DefaultDirName={localappdata}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=Output
OutputBaseFilename=CikkCheckerSetup_{#MyAppVersion}
SetupIconFile=assets\icon.ico
WizardSmallImageFile=assets\logo_small.png
Compression=lzma2/ultra
SolidCompression=yes
MinVersion=10.0
ArchitecturesInstallIn64BitMode=x64compatible
ArchitecturesAllowed=x64compatible
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
WizardStyle=modern
; Close running instances before installing to avoid DLL lock errors
CloseApplications=yes
CloseApplicationsFilter=CikkChecker.exe
RestartApplications=no
UninstallDisplayName={#MyAppName}
UninstallDisplayIcon={app}\{#MyAppExeName}
ShowLanguageDialog=yes

[Languages]
Name: "hungarian";  MessagesFile: "compiler:Languages\Hungarian.isl"
Name: "english";    MessagesFile: "compiler:Default.isl"
Name: "german";     MessagesFile: "compiler:Languages\German.isl"
Name: "slovak";     MessagesFile: "compiler:Languages\Slovak.isl"
Name: "czech";      MessagesFile: "compiler:Languages\Czech.isl"
Name: "polish";     MessagesFile: "compiler:Languages\Polish.isl"
Name: "french";     MessagesFile: "compiler:Languages\French.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
Source: "CikkChecker_app.exe"; DestDir: "{app}"; DestName: "CikkChecker.exe"; Flags: ignoreversion
Source: "checker_core.exe";    DestDir: "{app}"; Flags: ignoreversion
Source: "assets\*";            DestDir: "{app}\assets"; Flags: ignoreversion recursesubdirs

[Icons]
Name: "{group}\{#MyAppName}";           Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}";     Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#MyAppName}}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{app}"
