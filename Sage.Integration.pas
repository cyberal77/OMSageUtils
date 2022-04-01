unit Sage.Integration;

interface

type
  TSageApplication = (saCial, saCpta, saMopa);

// Records for parameters
  TSageExternalProgExecutable = record
    Nom:        string;
    Contexte:   cardinal;
    Path:       string;
    Parametres: array of string;
    Attente:    boolean;
    Fermeture:  boolean;
  end;

function GetSageAppFullPath(const SageApplication: TSageApplication): string;
function GetSageAppVersion(const SageApplication: TSageApplication): string;

// Fonction add
function SageAddExternalProgramExecutable(const SageApplication: TSageApplication;
                                Parameters: TSageExternalProgExecutable): boolean;

const
  // Type
  SageProgramExterneExecutable = $66696c65;
  SageProgramExterneLienInternet = $75726c20;
  SageProgramExterneLienInternetIntegre = $756c6e6b;
  SageProgramExternePageWebIntegree = $666c6e6b;
//  SageProgramExterneScriptIntegre = ????

  // Contexte
  SageContexteGlobal              = $7d0;
  SageContexteTiers               = $7d1;
  SageContexteClients             = $7d2;
  SageContexteSectionsAnalytiques = $7d3;
  SageContexteBanques             = $7d4;
  SageContexteArticles            = $7d6;
  SageContexteDocumentsVente      = $7d7;
  SageContexteDocumentsAchat      = $7d8;
  SageContexteDocumentsStock      = $7d9;
  SageContexteDocumentsInterne    = $7da;
  SageContexteLignesDocument      = $7db;
  SageContexteCollaborateurs      = $7dc;
  SageContexteRessources          = $7dd;
  SageContexteDepots              = $7de;
  SageContexteProjetsFabrication  = $7df;
  SageContexteProjetsAffaire      = $7e0;
  SageContexteMail                = $7dd;
  SageContexteMailRecouvrement    = $7de;
  SageContexteEcrituresComptables = $10eb;

implementation

uses
  System.SysUtils,
  System.IOUtils,
  WinApi.Windows,
  Winapi.TlHelp32,
  Win.Registry,
  System.Classes;

const
  SageKeys: array [0..2] of string = ('Gestion Commerciale 100c', 'Comptabilité 100c', 'Moyens de paiement 100c');
  SageExes: array [0..2] of string = ('GecoMaes.exe', 'Maestria.exe', 'MopaMaes.exe');
  HKLMRootKey     = 'SOFTWARE\WOW6432Node\Sage\%s';
  HKCURootKey     = 'SOFTWARE\Sage\%s\%s';
  ExternalProgKey = 'Personnalisation\Programmes externes';
  StartProgId     = 61962;

// Retrive ProductInfoVersion


function GetPath(const SageApplication: TSageApplication): string;
var
  Reg: TRegistry;
  Key: string;
begin
  Result := '';
  Reg := TRegistry.Create();
  try
    Reg.RootKey := HKEY_LOCAL_MACHINE;
    Key := Format(HKLMRootKey, [SageKeys[Ord(SageApplication)]]);
    if Reg.KeyExists(Key) then begin
      Reg.OpenKeyReadOnly(key);
      if Reg.ValueExists('Path') then
        Result := Reg.ReadString('Path');
    end;
  finally
    Reg.Free;
  end;
end;

function GetSageAppFullPath(const SageApplication: TSageApplication): string;
begin
  Result := TPath.Combine(GetPath(SageApplication), SageExes[Ord(SageApplication)]);
  if (not FileExists(Result)) then Result := '';
end;

function GetVersion(const FileName: string): string;
var
  VerInfoSize, {VerValueSize,} Dummy: DWord;
  VerInfo, VerString: Pointer;
  {VerValue: PVSFixedFileInfo;}
const
  Lang = '040904E4'; // Default Lang Code for Neutral?
begin
  if (FileName <> '') then begin
    VerInfoSize := GetFileVersionInfoSize(PChar(FileName), Dummy);
    if VerInfoSize <> 0 then
    {- Les info de version sont inclues }
    begin
      {On alloue de la mémoire pour un pointeur sur les info de version : }
      GetMem(VerInfo, VerInfoSize);
      {On récupère l'information : }
      GetFileVersionInfo(PChar(FileName), 0, VerInfoSize, VerInfo);
      if VerQueryValue(VerInfo, PChar('\StringFileInfo\' + lang + '\ProductVersion'), Pointer(VerString), VerInfoSize) then
        Result := PChar(VerString)
      else
        Result := '';

      {On libère la place précédemment allouée : }
      FreeMem(VerInfo, VerInfoSize);
    end else
      result := '';
  end;
end;

function GetSageAppVersion(const SageApplication: TSageApplication): string;
begin
  Result := GetVersion(GetSageAppFullPath(SageApplication));
end;

// Function for add external Program
function GetId(const ProgKey: string): integer;
var
  Idx: integer;
  Reg: TRegistry;
begin
  Reg := TRegistry.Create();
  Result := -1;
  try
    if (Reg.OpenKey(ProgKey, True)) then begin
      Idx := StartProgId;
      while (Reg.KeyExists(Idx.ToString)) do Inc(Idx);
      Result := Idx;
    end;
  finally
    Reg.Free;
  end;
end;

// Fonction to check if prog already exist with same path and context
function ProgIsSet(const ProgKey, Nom: string; const Contexte: Integer): integer;
var
  Reg: TRegistry;
  ProgList: TStringList;
  Idx: Integer;
begin
  Reg := TRegistry.Create;
  ProgList := TStringList.Create();
  Result := 0;
  try
    if Reg.KeyExists(ProgKey) then begin
      Reg.OpenKeyReadOnly(ProgKey);
      Reg.GetKeyNames(ProgList);
      Reg.CloseKey;
      Idx := 0;
      while ((Idx < ProgList.Count) and (Result = 0)) do begin
        Reg.OpenKeyReadOnly(TPath.Combine(ProgKey, ProgList.Strings[Idx]));
        if (Reg.ValueExists('Nom') and Reg.ValueExists('Contexte')) then begin
          if ((CompareStr(Reg.ReadString('Nom'), Nom) = 0) and
             (Reg.ReadInteger('Contexte') = Contexte)) then
               Result := ProgList.Strings[Idx].ToInteger;
        end;
        Reg.CloseKey;
        Inc(Idx);
      end;
    end;
  finally
    Reg.Free;
    ProgList.Free;
  end;
end;

function CheckIfSageAppRunning(const Path: string): boolean;
var
  ContinueLoop: BOOL;
  FSnapshotHandle: THandle;
  FProcessEntry32: TProcessEntry32;
  ExeFileName, ProcessFileName: string;
begin
  ExeFileName := TPath.GetFileName(Path);
  FSnapshotHandle := CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
  FProcessEntry32.dwSize := SizeOf(FProcessEntry32);
  ContinueLoop := Process32First(FSnapshotHandle, FProcessEntry32);
  Result := False;
  while Integer(ContinueLoop) <> 0 do
  begin
    ProcessFileName := TPath.GetFileName(FProcessEntry32.szExeFile);
    if (CompareStr(ExeFileName, ProcessFileName)= 0) then
    begin
      Result := True;
    end;
    ContinueLoop := Process32Next(FSnapshotHandle, FProcessEntry32);
  end;
  CloseHandle(FSnapshotHandle);
end;

function SageAddExternalProgramExecutable(const SageApplication: TSageApplication;
                                Parameters: TSageExternalProgExecutable): boolean;
var
  Reg: TRegistry;
  ProgIdx: integer;
  ProgKey: string;
  ParamStr: string;
  Idx: Integer;
begin
  Result := False;
  Reg := TRegistry.Create();
  if (not CheckIfSageAppRunning(GetSageAppFullPath(SageApplication))) then begin
    ProgKey := Format(HKCURootKey, [SageKeys[Ord(SageApplication)], GetSageAppVersion(SageApplication)]);
    ProgKey := TPath.Combine(ProgKey, ExternalProgKey);
    try
      ProgIdx := ProgIsSet(ProgKey, Parameters.Nom, Parameters.Contexte);
      if (ProgIdx = 0) then ProgIdx := GetId(ProgKey);
      // Write Values
      if Reg.OpenKey(TPath.Combine(ProgKey, ProgIdx.ToString), True) then begin
        Reg.WriteInteger('Type', SageProgramExterneExecutable);
        Reg.WriteString('Nom', Parameters.Nom);
        Reg.WriteInteger('Contexte', Parameters.Contexte);
        Reg.WriteString('Path', Parameters.Path);
        // Array to string
        ParamStr := '';
        for Idx := Low(Parameters.Parametres) to High(Parameters.Parametres) do begin
          ParamStr := ParamStr + Parameters.Parametres[Idx] + ' ';
        end;
        Reg.WriteString('Paramètres', Trim(ParamStr));
        if Parameters.Attente then Reg.WriteString('Attente', '1')
        else Reg.WriteString('Attente', '0');
        if Parameters.Fermeture then Reg.WriteString('Fermeture société', '1')
        else Reg.WriteString('Fermeture société', '0');
      end;
    finally
      Reg.Free;
    end;
  end else
    raise Exception.Create('Le programme Sage est en cours d''execution');
end;


end.
