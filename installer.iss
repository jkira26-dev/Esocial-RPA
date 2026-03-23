; ============================================================
;  eSocial RPA — Script de Instalação (Inno Setup 6)
;  Para compilar: abra este arquivo no Inno Setup Compiler
;  Download: https://jrsoftware.org/isdl.php
; ============================================================

#define AppName    "eSocial RPA"
#define AppVersion "1.0.0"
#define AppPublisher "IOB Gestão Contábil"
#define AppExeName "eSocialRPA.exe"
#define AppDir     "eSocialRPA"

[Setup]
AppId={{8F3A2C1D-4B7E-4F9A-BC12-D5E6F7A8B9C0}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisherURL=https://www.iob.com.br
AppSupportURL=https://www.iob.com.br
AppUpdatesURL=https://www.iob.com.br
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
AllowNoIcons=yes
LicenseFile=
OutputDir=installer_output
OutputBaseFilename=eSocialRPA_Instalador_v{#AppVersion}
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
DisableProgramGroupPage=yes
PrivilegesRequired=admin
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64
SetupIconFile=icon.ico
UninstallDisplayIcon={app}\{#AppExeName}
MinVersion=10.0
CloseApplications=yes

[Languages]
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Tasks]
Name: "desktopicon";  Description: "Criar ícone na área de trabalho"; \
      GroupDescription: "Ícones adicionais:"
Name: "quicklaunch"; Description: "Fixar na barra de tarefas"; \
      GroupDescription: "Ícones adicionais:"; Flags: unchecked

[Files]
; Binário gerado pelo PyInstaller (pasta dist\eSocialRPA\)
Source: "dist\{#AppDir}\*"; DestDir: "{app}"; \
        Flags: ignoreversion recursesubdirs createallsubdirs

; Arquivos de configuração editáveis — na pasta de dados do usuário
Source: "config.py";  DestDir: "{userappdata}\eSocialRPA"; \
        Flags: onlyifdoesntexist
Source: "README.md";  DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#AppName}";           Filename: "{app}\{#AppExeName}"
Name: "{group}\Desinstalar {#AppName}"; Filename: "{uninstallexe}"
Name: "{commondesktop}\{#AppName}";   Filename: "{app}\{#AppExeName}"; \
      Tasks: desktopicon

[Run]
; Instala dependências Python necessárias para notificações (opcional)
Filename: "{cmd}"; Parameters: "/C pip install plyer --quiet"; \
          StatusMsg: "Instalando suporte a notificações..."; \
          Flags: runhidden waituntilterminated; Check: PythonInstalado

; Abre README após instalação
Filename: "{app}\README.md"; Description: "Ver instruções de uso"; \
          Flags: postinstall skipifsilent shellexec

[UninstallDelete]
Type: filesandordirs; Name: "{userappdata}\eSocialRPA\*.json"
Type: filesandordirs; Name: "{userappdata}\eSocialRPA\*.log"

[Messages]
BeveledLabel={#AppName} v{#AppVersion}
WelcomeLabel1=Bem-vindo ao instalador do {#AppName}
WelcomeLabel2=Este programa instala o {#AppName} — download automático de XMLs do eSocial.%n%nRequisitos:%n  • Windows 10 ou superior (64-bit)%n  • Google Chrome instalado%n  • Conexão com a internet%n%nClique em Avançar para continuar.
FinishedLabel=O {#AppName} foi instalado com sucesso.%n%nPara iniciar, clique em {#AppExeName} na área de trabalho ou no menu Iniciar.%n%nIMPORTANTE: Antes de usar, execute o 2_ABRIR_CHROME.bat para preparar o Chrome com a porta de debug.
FinishedHeadingLabel=Instalação concluída!

[Code]
{ Verifica se o Python está instalado — usado para instalar plyer }
function PythonInstalado: Boolean;
var
  ResultCode: Integer;
begin
  Result := Exec('python', '--version', '', SW_HIDE, ewWaitUntilTerminated, ResultCode)
            and (ResultCode = 0);
end;

{ Verifica se o Chrome está instalado }
function ChromeInstalado: Boolean;
begin
  Result := FileExists('C:\Program Files\Google\Chrome\Application\chrome.exe')
         or FileExists('C:\Program Files (x86)\Google\Chrome\Application\chrome.exe');
end;

{ Avisa sobre Chrome ausente }
procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    if not ChromeInstalado then
      MsgBox('Atenção: Google Chrome não foi encontrado.' + #13#10 +
             'O eSocial RPA requer o Chrome instalado para funcionar.' + #13#10 +
             'Instale o Chrome em: https://www.google.com/chrome',
             mbInformation, MB_OK);
  end;
end;
