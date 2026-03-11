; Inno Setup script for Clipboard Coach
; Download Inno Setup from https://jrsoftware.org/isdl.php to compile this

[Setup]
AppName=Clipboard Coach
AppVersion=1.0
AppPublisher=Clipboard Coach
DefaultDirName={autopf}\ClipboardCoach
DefaultGroupName=Clipboard Coach
OutputBaseFilename=ClipboardCoach-Setup
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=lowest
SetupIconFile=compiler:SetupClassicIcon.ico

[Files]
Source: "dist\ClipboardCoach.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "config.example.json"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Clipboard Coach"; Filename: "{app}\ClipboardCoach.exe"
Name: "{group}\Uninstall Clipboard Coach"; Filename: "{uninstallexe}"
Name: "{userstartup}\Clipboard Coach"; Filename: "{app}\ClipboardCoach.exe"; Comment: "Start Clipboard Coach at login"

[Run]
Filename: "{app}\ClipboardCoach.exe"; Description: "Launch Clipboard Coach"; Flags: postinstall nowait skipifsilent
