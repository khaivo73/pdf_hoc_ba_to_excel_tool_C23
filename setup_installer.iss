; Inno Setup Script - HocBaC23PdfToExcel
; Tac gia: Vo Thanh Khai

#define AppName "HocBaC23PdfToExcel"
#define AppNameVN "Chuyen PDF Hoc Ba sang Excel"
#define AppVersion "1.1.0"
#define AppPublisher "Vo Thanh Khai"
#define AppExeName "HocBaC23PdfToExcel.exe"
#define AppURL "https://github.com/khaivo73/pdf_hoc_ba_to_excel_tool_C23"

[Setup]
AppId={{A3F2B8C1-4D7E-4F9A-B2C3-D8E5F1A6B7C4}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppNameVN} v{#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}
AppUpdatesURL={#AppURL}/releases
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
OutputDir=dist
OutputBaseFilename=HocBaC23PdfToExcel_v{#AppVersion}_Setup
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequiredOverridesAllowed=dialog
UninstallDisplayIcon={app}\{#AppExeName}
VersionInfoVersion={#AppVersion}
VersionInfoCompany={#AppPublisher}
VersionInfoDescription={#AppNameVN}
ShowLanguageDialog=no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Tao bieu tuong tren man hinh (Desktop)"; GroupDescription: "Bieu tuong:"; Flags: unchecked
Name: "quicklaunchicon"; Description: "Tao bieu tuong tren thanh Taskbar"; GroupDescription: "Bieu tuong:"; Flags: unchecked; OnlyBelowVersion: 6.1

[Files]
Source: "dist\{#AppExeName}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#AppNameVN}"; Filename: "{app}\{#AppExeName}"
Name: "{group}\Gỡ cài đặt {#AppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppNameVN}"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: quicklaunchicon

[Run]
Filename: "{app}\{#AppExeName}"; Description: "Chay ung dung ngay bay gio"; Flags: nowait postinstall skipifsilent
