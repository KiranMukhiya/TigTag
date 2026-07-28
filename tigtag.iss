; Inno Setup script for TigTag.
;
; Build order:
;   1. pyinstaller tigtag.spec          -> produces dist\TigTag.exe
;   2. Compile this script with Inno Setup (ISCC tigtag.iss, or open in the
;      Inno Setup IDE and click Compile)
;      -> produces installer_output\TigTag_Setup_v1.0.exe
;
; Requires tigtag.ico (multi-size .ico, generated from tigtag.png) to sit
; next to this script and next to tigtag.py.

#define MyAppName "TigTag"
#define MyAppVersion "1.0"
#define MyAppPublisher "Kiran Mukhiya"
#define MyAppContact "mailto:kiranmkya@gmail.com"
#define MyAppExeName "TigTag.exe"

[Setup]
AppId={{80080163-1080-4F97-88D3-1B1521EF10EC}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppContact}
AppSupportURL={#MyAppContact}
AppContact={#MyAppContact}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
LicenseFile=LICENSE
OutputDir=installer_output
OutputBaseFilename=TigTag_Setup_v{#MyAppVersion}
SetupIconFile=tigtag.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequiredOverridesAllowed=dialog

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; The onefile PyInstaller build already has tigtag.png bundled inside the
; exe, so only the exe itself (plus docs) needs to ship here.
Source: "dist\TigTag.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "README.md"; DestDir: "{app}"; Flags: ignoreversion isreadme
Source: "LICENSE"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Ask the user on uninstall whether to keep locally-saved config/feedback;
; by default Inno leaves user data alone. Uncomment to remove it too:
; Type: filesandordirs; Name: "{userappdata}\..\.tigtag"
