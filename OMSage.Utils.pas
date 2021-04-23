unit OMSage.Utils;

interface

uses
  IniFiles, classes, {$INCLUDE OmSageLib.inc};

procedure CreateSageCialApplication(var ASCial: IBSCialApplication3; const AVersion: string = '');
procedure CreateSageCptaApplication(var ASCpta: IBSCptaApplication3; const AVersion: string = '');

function ConnectSageCial(AIni: TIniFile; var ASCial: IBSCialApplication3;
  const ASection: string = 'Sage'; const AUserName: string = ''; const APassword: string = ''): boolean; overload;

function ConnectSageCial(var ASCial: IBSCialApplication3;
  const AFileName: string = 'C:\Users\Public\Documents\Sage\iGestion commerciale\Bijou.gcm';
  const AUserName: string = '<Administrateur>'; const APassword: string = ''): boolean; overload;

function ConnectSageCpta(AIni: TIniFile; var ASCpta: IBSCptaApplication3;
const ASection: string = 'Sage'; const AUserName: string = ''; const APassword: string = ''): boolean; overload;

function ConnectSageCpta(var ASCpta: IBSCptaApplication3;
  const AFileName: string = 'C:\Users\Public\Documents\Sage\iGestion commerciale\Bijou.mae';
  const AUserName: string = '<Administrateur>'; const APassword: string = ''): boolean; overload;

procedure DisconnectSageOM(ASCial: IBSCialApplication3); overload;
procedure DisconnectSageOM(ASCpta: IBSCptaApplication3); overload;

// Quelque fonction standard pratique
function DeleteDocIfEmpty(ADoc: IBODocument3): integer;
function getCANum(ACpta: IBSCptaApplication3; ACANum: string;
  APlan: string = 'Plan Affaire'): IBOCompteA3;
function getTiersStrings(ACpta: IBSCptaApplication3;
  ATypeTiers: TiersType; Strings: TStrings; const All: boolean = False;
  AFrom: string = ''; ATo: string = ''): integer;
function getCompteAStrings(ACpta: IBSCptaApplication3; Strings: TStrings;
  const APlanAnalytique: string = ''; const All: boolean = False): integer;
function getPlanAnalytiqueAffaire(ASCial: IBSCialApplication3): string;
function getPlanAnalytiqueArticle(ASCial: IBSCialApplication3): string;
function getArticleStrings(ACial: IBSCIalApplication3;
  Strings: TStrings; const All: boolean = False;
  AFrom: string = ''; ATo: string = ''): integer;
function getArticle(ACial: IBSCIalApplication3; const AR_Ref: string): IBOArticle3;

// Fonction Compta
function getCurrentExercice(ACpta: IBSCptaApplication3; var DateDeb, DateFin: TDate): integer;

implementation

uses
  SysUtils, UCipher, ComObj, Forms;

function getOMContext(const AVersion: string): OleVariant;
var
  ManifestFileName: string;
begin
  ManifestFileName := Format('%s\OM\Objets100c-v%s.manifest',[ExtractFilePath(Application.ExeName), AVersion]);
  try
    Result := CreateOleObject('Microsoft.Windows.ActCtx');
    if FileExists(ManifestFileName) then 
      Result.Manifest := ManifestFileName;
  except
    Result := false;
  end;
end;

procedure CreateSageCialApplication(var ASCial: IBSCialApplication3; const AVersion: string);
begin
  if ASCial = nil then begin
    try
    {$IFDEF SAGE100}
      ASCial := IUnKnown(getOMContext(AVersion).CreateObject(ClassIDToProgID(CLASS_BSCIALApplication3))) as IBSCIALApplication3;
    {$ELSE}
      ASCial := IUnKnown(getOMContext(AVersion).CreateObject(ClassIDToProgID(CLASS_BSCIALApplication100c))) as IBSCIALApplication3;
    {$ENDIF}
    except
      ASCial := nil;
    end;
  end;
end;

procedure CreateSageCptaApplication(var ASCpta: IBSCptaApplication3; const AVersion: string);
begin
  if ASCpta = nil then begin
  {$IFDEF SAGE100}
    ASCpta := IUnKnown(getOMContext(AVersion).CreateObject(ClassIDToProgID(CLASS_BSCPTAApplication3))) as IBSCPTAApplication3;
  {$ELSE}
    ASCpta := IUnKnown(getOMContext(AVersion).CreateObject(ClassIDToProgID(CLASS_BSCPTAApplication100c))) as IBSCPTAApplication3;
  {$ENDIF}
  end;
end;

function ConnectSageCial(AIni: TIniFile; var ASCial: IBSCialApplication3;
  const ASection, AUsername, APassword: string): boolean;
var
  GescoFileName: string;
  Username, Password: string;
begin
  try
    // Get From Ini
    GescoFileName := AIni.ReadString(ASection,'CialFileName','C:\Users\Public\Documents\Sage\iGestion commerciale\Bijou.gcm');
    if (AUsername = '') then begin
      Username := AIni.ReadString(ASection,'Username','<Administrateur>');
      Password := XORDecode(AIni.ReadString(ASection,'Password',''));
    end else begin
      Username := AUsername;
      Password := APassword;
    end;
    Result := ConnectSageCial(ASCial, GescoFileName, Username, Password);
  except
    Result := False;
  end;
end;

function ConnectSageCpta(AIni: TIniFile; var ASCpta: IBSCptaApplication3;
  const ASection, AUsername, APassword: string): boolean;
var
  CptaFileName: string;
  Username, Password: string;
begin
  try
    // Get From Ini
    CptaFileName := AIni.ReadString(ASection,'CptaFileName','C:\Users\Public\Documents\Sage\iGestion commerciale\Bijou.mae');
    if (AUsername = '') then begin
      Username := AIni.ReadString(ASection,'Username','<Administrateur>');
      Password := XORDecode(AIni.ReadString(ASection,'Password',''));
    end else begin
      Username := AUsername;
      Password := APassword;
    end;
    //Configure Object100
    Result := ConnectSageCpta(ASCpta, CptaFileName, Username, Password);
  except
    Result := False;
  end;
end;

function ConnectSageCial(var ASCial: IBSCialApplication3;
  const AFileName, AUserName, APassword: string): boolean;
var
  GescoFileName: string;
begin
  try
    DisconnectSageOM(ASCial);
    CreateSageCialApplication(ASCial);
    // Get From Ini
    GescoFileName := AFileName;
    //Configure Object100
    ASCial.Name := GescoFileName;
    ASCial.Loggable.UserName := AUsername;
    ASCial.Loggable.UserPwd := APassword;
    ASCial.Open;
    Result := ASCial.IsOpen;
  except
    Result := False;
  end;
end;

function ConnectSageCpta(var ASCpta: IBSCptaApplication3;
  const AFileName, AUserName, APassword: string): boolean;
var
  CptaFileName: string;
begin
  try
    CreateSageCptaApplication(ASCpta);
    // Get From Ini
    CptaFileName := AFileName;
    //Configure Object100
    ASCpta.Name := CptaFileName;
    ASCpta.Loggable.UserName := AUsername;
    ASCpta.Loggable.UserPwd := APassword;
    ASCpta.Open;
    Result := ASCpta.IsOpen;
  except
    Result := False;
  end;
end;

procedure DisconnectSageOM(ASCial: IBSCialApplication3);
begin
  if (ASCial <> nil) then
    if ASCial.IsOpen then ASCial.Close;
end;

procedure DisconnectSageOM(ASCpta: IBSCptaApplication3);
begin
  if (ASCpta <> nil) then
    if ASCpta.IsOpen then ASCpta.Close;
end;

// Fonctions Utiles

{*
  fonction DeleteDocIfEmpty
  Paramètre :
    ADoc: IBODocument3 : Un objet Document Entête
  Renvoie 0 si le document a bien été supprimé
          -1 si une erreur c'est produite
          >0 le nombre de lignes du document
*}
function DeleteDocIfEmpty(ADoc: IBODocument3): integer;
begin
  if ADoc.FactoryDocumentLigne.List.Count = 0 then begin
    try
      ADoc.Remove;
      Result := 0;
    except
      Result := -1;
    end;
  end
  else Result := ADoc.FactoryDocumentLigne.List.Count;
end;

function getCANum(ACpta: IBSCptaApplication3; ACANum: string;
  APlan: string = 'Plan Affaire'): IBOCompteA3;
var
  PlanAna: IBPAnalytique3;
begin
  // Analytique;
  try
    PlanAna := ACpta.FactoryAnalytique.ReadIntitule(APlan);
    if (ACANum <> '') then begin
      if (ACpta.FactoryCompteA.ExistNumero(PlanAna,ACANum)) then
        Result := ACpta.FactoryCompteA.ReadNumero(PlanAna, ACANum);
      end else Result := nil;
  except
    Result := nil;
  end;
end;

function getTiersStrings(ACpta: IBSCptaApplication3;
  ATypeTiers: TiersType; Strings: TStrings; const All: boolean = False;
  AFrom: string = ''; ATo: string = ''): integer;
var
  Coll: IBICollection;
  Idx: Integer;
begin
  if AFrom = '' then AFrom := '0';
  if ATo = '' then ATo := 'Z';
  Result := 0;
  Coll := ACpta.FactoryTiers.QueryTypeNumeroOrderNumero(ATypeTiers, AFrom, ATo);
  Strings.BeginUpdate;
  for Idx := 1 to Coll.Count do begin
    if ((not (Coll.Item[Idx] as IBOTiers3).CT_Sommeil) or (All)) then begin
      Strings.Add((Coll.Item[Idx] as IBOTiers3).CT_Num);
      Inc(Result);
    end;
  end;
  Strings.EndUpdate;
end;

function getCompteAStrings(ACpta: IBSCptaApplication3; Strings: TStrings;
  const APlanAnalytique: string = ''; const All: boolean = False): integer;
var
  Coll: IBICollection;
  Idx: Integer;
  Plan: IBPAnalytique3;
begin
  if (APlanAnalytique = '') then Exit(0)
  else begin
    if (not ACpta.FactoryAnalytique.ExistIntitule(APlanAnalytique)) then Exit(0)
    else Plan := ACpta.FactoryAnalytique.ReadIntitule(APlanAnalytique);
  end;
  Result := 0;
  if Plan <> nil then begin
    Coll := ACpta.FactoryCompteA.QueryPlanA(Plan);
    Strings.BeginUpdate;
    for Idx := 1 to Coll.Count do begin
      if ((not (Coll.Item[Idx] as IBOCompteA3).CA_Sommeil) or (All)) then begin
        Strings.Add((Coll.Item[Idx] as IBOCompteA3).CA_Num);
        Inc(Result);
      end;
    end;
    Strings.EndUpdate;
  end;
end;

function getPlanAnalytiqueAffaire(ASCial: IBSCialApplication3): string;
var
  Analytique: IBPAnalytique3;
begin
  Analytique := (ASCial.FactoryDossierParam.List.Item[1] as IBPDossierParamCial).AnalytiqueAffaire;
  if Analytique <> nil then Result := Analytique.A_Intitule
  else Result := '';
end;

function getPlanAnalytiqueArticle(ASCial: IBSCialApplication3): string;
var
  Analytique: IBPAnalytique3;
begin
  Analytique := (ASCial.FactoryDossierParam.List.Item[1] as IBPDossierParamCial).AnalytiqueArticle;
  if Analytique <> nil then Result := Analytique.A_Intitule
  else Result := '';
end;

function getArticleStrings(ACial: IBSCIalApplication3;
  Strings: TStrings; const All: boolean = False;
  AFrom: string = ''; ATo: string = ''): integer;
var
  Coll: IBICollection;
  Idx: Integer;
begin

  if All then begin
    if AFrom = '' then AFrom := '0';
    if ATo = '' then ATo := 'Z';
    Coll := ACial.FactoryArticle.QueryReferenceOrderReference(AFrom, ATo);
  end else begin
    if ((AFrom <> '') or (ATo <> '')) then begin
      if AFrom = '' then AFrom := '0';
      if ATo = '' then ATo := 'Z';
      Coll := ACial.FactoryArticle.QueryReferenceOrderReference(AFrom, ATo);
    end else Coll := ACial.FactoryArticle.QueryActifOrderReference();
  end;
  Result := 0;
  // Traitements
  Strings.BeginUpdate;
  for Idx := 1 to Coll.Count do begin
    if ((not (Coll.Item[Idx] as IBOArticle3).AR_Sommeil) or (All)) then begin
      Strings.Add((Coll.Item[Idx] as IBOArticle3).AR_Ref);
      Inc(Result);
    end;
  end;
  Strings.EndUpdate;
end;

function getArticle(ACial: IBSCIalApplication3; const AR_Ref: string): IBOArticle3;
begin
  if ACial.FactoryArticle.ExistReference(AR_Ref) then begin
    Result := ACial.FactoryArticle.ReadReference(AR_Ref);
  end else
    Result := nil;
end;


function getCurrentExercice(ACpta: IBSCptaApplication3; var DateDeb, DateFin: TDate): integer;
var
  C: boolean;
  Dossier: IBPDossier2;
begin
  C := True;
  Result := 1;

  Dossier := ACpta.FactoryDossier.List.Item[1] as IBPDossier2;

  while ((C) and (Result < 6)) do begin
    try
      DateDeb := Dossier.D_DebutExo[Result];
      DateFin := Dossier.D_FinExo[Result];
      Inc(Result);
    except
      Dec(Result);
      C := False;
    end;
  end;
end;

end.
