; ============================================================
; Inno Setup Script — Formulir Ritual Installer
; ============================================================
; Membundel aplikasi FormulirRitual.exe beserta semua
; dependensi dari folder dist\FormulirRitual ke installer .exe
; yang mudah diinstal di perangkat lain tanpa Python.
; ============================================================

#define MyAppName "Formulir Ritual"
#define MyAppVersion "1.3.0"
#define MyAppPublisher "Harum Semerbak"
#define MyAppExeName "FormulirRitual.exe"
#define MyAppURL "https://github.com/kebo-sukses/Harum-Semerbak-app"

[Setup]
; Info dasar aplikasi
AppId={{B3F7A1C2-9D4E-4F6B-8A1D-2E3C5F7A9B1D}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}

; Direktori install default
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}

; Output installer
OutputDir=installer_output
OutputBaseFilename=FormulirRitual_Setup_{#MyAppVersion}

; Kompresi
Compression=lzma2/ultra64
SolidCompression=yes
LZMAUseSeparateProcess=yes

; Tampilan
WizardStyle=modern
SetupIconFile=
DisableProgramGroupPage=yes

; Hak akses
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog

; Informasi versi
VersionInfoVersion={#MyAppVersion}
VersionInfoCompany={#MyAppPublisher}
VersionInfoProductName={#MyAppName}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Semua file dari dist\FormulirRitual (exe + dependensi + assets)
Source: "dist\FormulirRitual\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; Start Menu shortcut
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"
; Desktop shortcut (opsional)
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; Jalankan aplikasi setelah instalasi selesai
Filename: "{app}\{#MyAppExeName}"; Description: "Jalankan {#MyAppName}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Bersihkan database & output saat uninstall
Type: filesandordirs; Name: "{app}\database\ritual_forms.db"
Type: filesandordirs; Name: "{app}\output"
