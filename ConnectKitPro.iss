; Script generated for Connect Kit Pro
#define MyAppName "Connect Kit Pro"
#define MyAppVersion "1.1"
#define MyAppPublisher "Braden Yates"
#define MyAppExeName "Connect Kit Pro.exe"

[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside Inno Setup)
AppId={{5303CE2A-A6C0-4B29-A498-ABF5665CE0E7}
AppName=Connect Kit Pro
AppVersion=1.1
AppPublisher=Braden Yates
AppCopyright=Copyright (C) 2026 Braden Yates
DefaultDirName={autopf}\{#MyAppName}
DisableProgramGroupPage=yes
; Only allow running on Windows 10/11
MinVersion=10.0
; Icon for the installer file itself (optional, uses your .ico)
SetupIconFile=C:\Users\byates\OneDrive - OBM\Desktop\My Applications\Connect Kit Pro\icon.ico
OutputBaseFilename=Connect Kit Pro Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; copy everything inside the folder, including subfolders
Source: "C:\Users\byates\OneDrive - OBM\Desktop\My Applications\Connect Kit Pro\dist\Connect Kit Pro\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon; IconFilename: "{app}\{#MyAppExeName}"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#MyAppName}}"; Flags: nowait postinstall skipifsilent