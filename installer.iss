; ============================================================================
; Instalador del Cotizador INNOBA Colombia DMC (Inno Setup)
; La version se pasa desde publicar_version.py:  ISCC /DMyAppVersion=1.2.0 installer.iss
; ============================================================================
#ifndef MyAppVersion
  #define MyAppVersion "1.1.0"
#endif
#define MyAppName "Cotizador INNOBA"
#define MyAppPublisher "INNOBA Colombia DMC"
#define MyAppExeName "CotizadorInnoba.exe"
#define MyAppURL "https://github.com/felipeortizjllo7-del/SOFTWARE-cotizador"

[Setup]
; AppId fijo: NO cambiar entre versiones (asi las nuevas versiones actualizan la misma app)
AppId={{7F3A9C21-4E8B-4D5A-B2E1-9A6F0C3D18E4}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} v{#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
; Instalacion POR USUARIO (sin pedir administrador -> las actualizaciones se aplican solas)
PrivilegesRequired=lowest
DefaultDirName={localappdata}\Programs\CotizadorInnoba
DisableProgramGroupPage=yes
DisableDirPage=yes
UninstallDisplayIcon={app}\{#MyAppExeName}
UninstallDisplayName={#MyAppName} v{#MyAppVersion}
OutputDir=installer_output
OutputBaseFilename=CotizadorInnoba-Setup-{#MyAppVersion}
SetupIconFile=app.ico
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
; Cierra la app si esta abierta (necesario al actualizar) y la reabre al final
CloseApplications=yes
RestartApplications=no

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "Crear un acceso directo en el escritorio"; GroupDescription: "Accesos directos:"

[Files]
Source: "dist\CotizadorInnoba.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Abrir {#MyAppName} ahora"; Flags: nowait postinstall skipifsilent
