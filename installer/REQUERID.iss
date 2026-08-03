#define MyAppName "REQUERID — Requerimento de GERID"
#define MyAppVersion "3.4"
#define MyAppPublisher "Rodriguez & Sousa Advogados Associados"
#define MyAppExeName "REQUERID Launcher.exe"
#define SourceDir "C:\Users\Pichau\Desktop\Petição Inicial\CODE\CLAUDE CODE\RequerimentoGERID\REQUERID - Instalador"

[Setup]
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppVerName={#MyAppName} {#MyAppVersion}
DefaultDirName={localappdata}\REQUERID
DefaultGroupName=REQUERID
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir=output
OutputBaseFilename=REQUERID Setup v{#MyAppVersion}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
UninstallDisplayName={#MyAppName}
UninstallDisplayIcon={app}\{#MyAppExeName}
DisableWelcomePage=no

[Languages]
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Tasks]
Name: "desktopicon"; Description: "Criar um atalho na área de trabalho"; GroupDescription: "Atalhos adicionais:"

[Files]
Source: "{#SourceDir}\REQUERID.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceDir}\_internal\*"; DestDir: "{app}\_internal"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#SourceDir}\REQUERID Launcher.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceDir}\_internal_launcher\*"; DestDir: "{app}\_internal_launcher"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#SourceDir}\versao.txt"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autodesktop}\REQUERID"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[UninstallDelete]
Type: filesandordirs; Name: "{app}"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Abrir o REQUERID agora"; Flags: nowait postinstall skipifsilent
