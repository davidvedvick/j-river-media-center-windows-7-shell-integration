; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

[Setup]
AppName=Media Center Windows 7 Taskbar Plugin for J. River Media Center 16
AppVerName=Media Center Windows 7 Taskbar Plugin for J. River Media Center 16
AppPublisher=DeVed Computers
AppPublisherURL=www.devedcomputers.com
AppSupportURL=www.devedcomputers.com
AppUpdatesURL=www.devedcomputers.com
DefaultDirName={pf}\J River\Media Center 16\Plugins\MC_Jumpbar
DefaultGroupName=MC Jumpbar Plugin
DisableProgramGroupPage=yes
OutputBaseFilename=MC_Jumpbar
Compression=lzma
SolidCompression=yes
DirExistsWarning=No

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "Build Files\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Icons]
Name: "{group}\{cm:UninstallProgram, Plugin}"; Filename: "{uninstallexe}"

[Registry]
Root: HKLM; Subkey: "Software\J. River\Media Jukebox\Plugins\Interface\MC_jumpbar"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\J. River\Media Jukebox\Plugins\Interface\MC_jumpbar"; ValueType: dword; ValueName: "IVersion"; ValueData: "00000001"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\J. River\Media Jukebox\Plugins\Interface\MC_jumpbar"; ValueType: string; ValueName: "Company"; ValueData: "DeVed Computers"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\J. River\Media Jukebox\Plugins\Interface\MC_jumpbar"; ValueType: string; ValueName: "Version"; ValueData: ".8"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\J. River\Media Jukebox\Plugins\Interface\MC_jumpbar"; ValueType: string; ValueName: "URL"; ValueData: "www.devedcomputers.com"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\J. River\Media Jukebox\Plugins\Interface\MC_jumpbar"; ValueType: string; ValueName: "Copyright"; ValueData: "Copyright (c) 2009, David Vedvick."; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\J. River\Media Jukebox\Plugins\Interface\MC_jumpbar"; ValueType: dword; ValueName: "PluginMode"; ValueData: "00000001"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\J. River\Media Jukebox\Plugins\Interface\MC_jumpbar"; ValueType: string; ValueName: "ProdID"; ValueData: "MCPlugin.MC_jumpbar"; Flags: uninsdeletekey

[Run]
Filename: "{win}\Microsoft.NET\Framework\v2.0.50727\regasm"; Parameters: "/Codebase MC_jumpbar.dll";          WorkingDir: "{app}\"; StatusMsg: "Registering Plugin"; Flags:runhidden

[UninstallRun]
Filename: "{win}\Microsoft.NET\Framework\v2.0.50727\regasm"; Parameters: "/unregister MC_jumpbar.dll";          WorkingDir: "{app}\"; StatusMsg: "Registering Plugin"; Flags:runhidden
