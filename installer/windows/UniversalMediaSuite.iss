; Universal Media Suite - Inno Setup installer
; Compile dari akar repo: ISCC.exe installer\windows\UniversalMediaSuite.iss
; Prasyarat: hasil PyInstaller onedir ada di dist\UniversalMediaSuite\

#define MyAppName "Universal Media Suite"
#define MyAppExeName "UniversalMediaSuite.exe"
#define MyAppVersion "2.0.0"
#define MyAppPublisher "Zidan Herlangga"
#define MyAppURL "https://zidan-herlangga.github.io/universal-media-suite/"
#define MyAppId "{{8ED369C9-26D6-407B-AB16-44B79CB3E795}"

[Setup]
AppId={#MyAppId}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={userpf}\{#MyAppName}
DisableProgramGroupPage=no
OutputDir=..\..\dist
OutputBaseFilename=UniversalMediaSuite_Setup
SetupIconFile=..\..\assets\favicon.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=lowest
RestartApplications=no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "..\..\dist\UniversalMediaSuite\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent