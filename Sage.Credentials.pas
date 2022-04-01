unit Sage.Credentials;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, System.IniFiles;

type
  TfrmSageCredentials = class(TForm)
    btnOk: TButton;
    btnCancel: TButton;
    lblUsername: TLabel;
    edtUsername: TEdit;
    lblPassword: TLabel;
    edtPassword: TEdit;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnOkClick(Sender: TObject);
  private
    FPasswordInifile: TIniFile;
    FUsername: string;
    FServer: string;
    FDBName: string;
    FPassword: string;
    FFileName: string;
    procedure SetPassword(const DBName, ServerName, Username, Password: string);
    function CheckConnection(const Filename, Username, Password: string): boolean;
    function CheckSageFile(const FileName: string): boolean;
    { Déclarations privées }
  public
    { Déclarations publiques }
    function ShowModal(const Filename, Username: string): TModalResult; reintroduce;
    function GetPassword(const DBName, ServerName, Username: string): string; overload;
    function GetPassword(const Filename, Username: string): string; overload;
    property Username: string read FUsername;
    property Password: string read FPassword;
    property DBName: string read FDBName;
    property Server: string read FServer;
    property Filename: string read FFilename;
  end;

implementation

{$R *.dfm}

uses
  UMyUtils,
  {$Include OMSageLib.inc},
  OMSage.Utils,
  System.IOUtils,
  Execute.Win.CryptString;

const
  StrCaption = 'Identifiant société : %s';

{ TFrmSageCredentials }

procedure TfrmSageCredentials.btnOkClick(Sender: TObject);
begin
  if (not CheckConnection(FFilename, edtUsername.Text, edtPassword.Text)) then begin
    ModalResult := mrRetry;
  end else begin
    SetPassword(FDBName, FServer, edtUserName.Text, edtPassword.Text);
  end;
end;

function TfrmSageCredentials.CheckConnection(const Filename, Username, Password: string): boolean;
var
  Cial: IBSCIALApplication3;
  Cpta: IBSCptaApplication3;
begin
  if (TPath.GetExtension(FileName) = '.gcm') then begin
    CreateSageCialApplication(Cial);
    Result := ConnectSageCial(Cial, FileName, Username, Password);
    DisconnectSageOM(Cial);
  end else begin
    CreateSageCptaApplication(Cpta);
    Result := ConnectSageCpta(Cpta, FileName, Username, Password);
    DisconnectSageOM(Cial);
  end;
end;

function TfrmSageCredentials.CheckSageFile(const FileName: string): boolean;
var
  Ini: TInifile;
begin
  Result := False;
  if (FileExists(Filename)) then begin
    FFilename := FileName;
    Ini := TInifile.Create(Filename);
    try
      FServer   := Ini.ReadString('CBASE','ServeurSQL','');
      FDBName   := TPath.GetFileNameWithoutExtension(FileName);
      Result    := True;
    finally
      Ini.Free;
    end;
  end;
end;

procedure TfrmSageCredentials.FormCreate(Sender: TObject);
begin
  FPasswordInifile := TInifile.Create(GetUserIniFileName(Self.Handle,'Sage','Credentials'));
end;

procedure TfrmSageCredentials.FormDestroy(Sender: TObject);
begin
  FPasswordInifile.Free;
end;

procedure TfrmSageCredentials.FormShow(Sender: TObject);
begin
  Caption := Format('%s sur %s', [FDBName, FServer]);
  edtUsername.Text := FUsername;
end;

function TfrmSageCredentials.GetPassword(const Filename,
  Username: string): string;
begin
  CheckSageFile(FileName);
  result := GetPassword(FDBName, FServer, Username);
end;

function TfrmSageCredentials.GetPassword(const DBName, ServerName, Username: string): string;
var
  Section: string;
begin
  Section := Format('%s:%s', [ServerName, DBName]);
  if (FPasswordInifile.SectionExists(Section)) then begin
    Result := FPasswordInifile.ReadString(Section, Username, '');
    if (Result <> '') then Result := DecryptStringBase64(Result);
  end else
    Result := '';
end;

procedure TfrmSageCredentials.SetPassword(const DBName, ServerName, Username, Password: string);
var
  Section: string;
  CryptedPassword: string;
begin
  Section := Format('%s:%s', [ServerName, DBName]);
  CryptedPassword := CryptStringBase64(Password);
  FPasswordInifile.WriteString(Section, Username, CryptedPassword);
  FPassword := Password;
end;

function TfrmSageCredentials.ShowModal(const Filename, Username: string): TModalResult;
begin
  if (CheckSageFile(Filename)) then begin
    FUsername := Username;
    Caption := Format(StrCaption, [FDBName]);
    Result := inherited ShowModal();
  end else
    Result := mrAbort;
end;

end.
