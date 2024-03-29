unit OMSage.Utils;

interface

uses
  System.IniFiles, System.classes, WinApi.ActiveX, {$INCLUDE OmSageLib.inc};

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

// Converts
function getSageDbTypeDocument(const OMTypeDocument: TOleEnum): integer;
function getOMTypeDocument(const SageDbTypeDocument: integer): TOleEnum;

implementation

uses
  SysUtils, Execute.Win.CryptString, ComObj;

function getOMContext(const AVersion: string): OleVariant;
var
  ManifestFileName: string;
begin
  ManifestFileName := Format('%s\OM\Objets100c-v%s.manifest',[ExtractFilePath(ParamStr(0)), AVersion]);
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
      Password := DecryptStringBase64(AIni.ReadString(ASection,'Password',''));
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
      Password := DecryptStringBase64(AIni.ReadString(ASection,'Password',''));
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
  Param�tre :
    ADoc: IBODocument3 : Un objet Document Ent�te
  Renvoie 0 si le document a bien �t� supprim�
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
    ACANum := UpperCase(ACANum);
    PlanAna := ACpta.FactoryAnalytique.ReadIntitule(APlan);
    if (ACANum <> '') then begin
      if (ACpta.FactoryCompteA.ExistNumero(PlanAna,ACANum)) then
        Result := ACpta.FactoryCompteA.ReadNumero(PlanAna, ACANum)
      else
        Result := nil;
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
  if ACial.FactoryArticle.ExistReference(UpperCase(AR_Ref)) then begin
    Result := ACial.FactoryArticle.ReadReference(UpperCase(AR_Ref));
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

// Converts
function getSageDbTypeDocument(const OMTypeDocument: TOleEnum): integer;
begin
  case OMTypeDocument of
    // Ventes
    DocumentTypeVenteDevis:             Result := 0;
    DocumentTypeVenteCommande:          Result := 1;
    DocumentTypeVentePrepaLivraison:    Result := 2;
    DocumentTypeVenteLivraison:         Result := 3;
    DocumentTypeVenteReprise:           Result := 4;
    DocumentTypeVenteAvoir:             Result := 5;
    DocumentTypeVenteFacture:           Result := 6;
    DocumentTypeVenteFactureCpta:       Result := 7;
    DocumentTypeVenteArchive:           Result := 8;
    // Achats
    DocumentTypeAchatDemande:           Result := 10;
    DocumentTypeAchatCommande:          Result := 11;
    DocumentTypeAchatCommandeConf:      Result := 12;
    DocumentTypeAchatLivraison:         Result := 13;
    DocumentTypeAchatReprise:           Result := 14;
    DocumentTypeAchatAvoir:             Result := 15;
    DocumentTypeAchatFacture:           Result := 16;
    DocumentTypeAchatFactureCpta:       Result := 17;
    DocumentTypeAchatArchive:           Result := 18;
    // Stocks
    DocumentTypeStockMouvIn:            Result := 20;
    DocumentTypeStockMouvOut:           Result := 21;
    DocumentTypeStockDeprec:            Result := 22;
    DocumentTypeStockVirement:          Result := 23;
    DocumentTypeStockPreparation:       Result := 24;
    DocumentTypeStockOrdreFabrication:  Result := 25;
    DocumentTypeStockFabrication:       Result := 26;
    DocumentTypeStockArchive:           Result := 27;
    // Internes
    DocumentTypeInterne1:               Result := 40;
    DocumentTypeInterne2:               Result := 41;
    DocumentTypeInterne3:               Result := 42;
    DocumentTypeInterne4:               Result := 43;
    DocumentTypeInterne5:               Result := 44;
    DocumentTypeInterne6:               Result := 45;
    DocumentTypeInterneArchive:         Result := 46;
    DocumentTypeInterne7:               Result := 47;
  else
    raise Exception.Create('Le type de document pass� en param�tre n''existe pas');
  end;
end;

function getOMTypeDocument(const SageDbTypeDocument: integer): TOleEnum;
begin
  case SageDbTypeDocument of
    // Ventes
    0:  Result := DocumentTypeVenteDevis;
    1:  Result := DocumentTypeVenteCommande;
    2:  Result := DocumentTypeVentePrepaLivraison;
    3:  Result := DocumentTypeVenteLivraison;
    4:  Result := DocumentTypeVenteReprise;
    5:  Result := DocumentTypeVenteAvoir;
    6:  Result := DocumentTypeVenteFacture;
    7:  Result := DocumentTypeVenteFactureCpta;
    8:  Result := DocumentTypeVenteArchive;
    // Achats
    10: Result := DocumentTypeAchatDemande;
    11: Result := DocumentTypeAchatCommande;
    12: Result := DocumentTypeAchatCommandeConf;
    13: Result := DocumentTypeAchatLivraison;
    14: Result := DocumentTypeAchatReprise;
    15: Result := DocumentTypeAchatAvoir;
    16: Result := DocumentTypeAchatFacture;
    17: Result := DocumentTypeAchatFactureCpta;
    18: Result := DocumentTypeAchatArchive;
    // Stocks
    20: Result := DocumentTypeStockMouvIn;
    21: Result := DocumentTypeStockMouvOut;
    22: Result := DocumentTypeStockDeprec;
    23: Result := DocumentTypeStockVirement;
    24: Result := DocumentTypeStockPreparation;
    25: Result := DocumentTypeStockOrdreFabrication;
    26: Result := DocumentTypeStockFabrication;
    27: Result := DocumentTypeStockArchive;
    // Internes
    40: Result := DocumentTypeInterne1;
    41: Result := DocumentTypeInterne2;
    42: Result := DocumentTypeInterne3;
    43: Result := DocumentTypeInterne4;
    44: Result := DocumentTypeInterne5;
    45: Result := DocumentTypeInterne6;
    46: Result := DocumentTypeInterneArchive;
    47: Result := DocumentTypeInterne7;
  else
    raise Exception.Create('Le type de document pass� en param�tre n''existe pas');
  end;
end;

end.
