; Inno Setup script for ClipFix
; Download Inno Setup from https://jrsoftware.org/isdl.php to compile this

[Setup]
AppName=ClipFix
AppVersion=1.0
AppPublisher=ClipFix
DefaultDirName={autopf}\ClipFix
DefaultGroupName=ClipFix
OutputBaseFilename=ClipFix-Setup
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=lowest
SetupIconFile=compiler:SetupClassicIcon.ico

[Files]
Source: "dist\ClipFix.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "config.example.json"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\ClipFix"; Filename: "{app}\ClipFix.exe"
Name: "{group}\Uninstall ClipFix"; Filename: "{uninstallexe}"
Name: "{userstartup}\ClipFix"; Filename: "{app}\ClipFix.exe"; Comment: "Start ClipFix at login"

[Run]
Filename: "{app}\ClipFix.exe"; Description: "Launch ClipFix"; Flags: postinstall nowait skipifsilent
