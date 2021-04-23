unit OMSage.Compta.Journaux.Verrou;

{*

Classe delphi depuis le code cs deThierry JOUVE - Sage Service Consulting

*}

interface

uses
  System.SysUtils, System.Classes, {$INCLUDE OmSageLib.inc};

type

  TJournalVerrou = class
  private
    FOmJournal: IBOJournal3;
    FPeriode: TDateTime;
    FOmCompteGeneral: IBOCompteG3;
    FOmEcriture: IBOEcriture3;
    FIsEcritureFictive: boolean;
    FIsJournalVerrouille: boolean;
    function GetNumPieceVerrou: string;
    function getPeriode: string;
  public
    constructor Create(AOmJournal: IBOJournal3; APeriode: TDateTime; AOmCompteGeneral: IBOCompteG3); reintroduce;
    destructor Destroy(); override;
    procedure Verrouiller();
    procedure Deverrouiller();
    property Periode: string read getPeriode;
    property NumPieceVerrou: string read GetNumPieceVerrou;
    property IsEcritureFictive: boolean read FIsEcritureFictive;
    property IsJournalVerrouille: boolean read FIsJournalVerrouille;
  end;

implementation

{ TJournalVerrou }

/// <summary>
/// Initialise un objet verrou sur un Journal et une periode.
/// </summary>
/// <param name="AOmJournal">Objet journal initialisé</param>
/// <param name="APeriode">Date periode</param>
/// <param name="AOmCompteGeneral">Objet Compte Général initialisé</param>
constructor TJournalVerrou.Create(AOmJournal: IBOJournal3; APeriode: TDateTime;
  AOmCompteGeneral: IBOCompteG3);
begin
  if AOmJournal <> nil then
    FOmJournal := AOmJournal
  else
    raise Exception.Create('AOmJournal n''est pas initialisé');
  FPeriode := APeriode;
  if AOmCompteGeneral <> nil then
    FOmCompteGeneral := AOmCompteGeneral
  else
    raise Exception.Create('AOmCompteGeneral n''est pas initialisé');
  FIsJournalVerrouille := False;
end;

destructor TJournalVerrou.Destroy;
begin
  if FIsJournalVerrouille then Deverrouiller();
  inherited;
end;


/// <summary>
/// Déverrouille la période du journal.
/// </summary>
procedure TJournalVerrou.Deverrouiller;
begin
  try
    if (FIsJournalVerrouille) then begin
      if (FIsEcritureFictive) then
        FOmEcriture.Remove()
      else
        FOmEcriture.Read();

      FIsJournalVerrouille := False;
    end;
  except
    on E: Exception do Raise Exception.CreateFmt('Erreur en déverrouillage du journal %s de %s : %s',
        [FOmJournal.JO_Num, FormatDateTime('mmmm yyyy', FPeriode), E.Message]);
  end;
end;

function TJournalVerrou.GetNumPieceVerrou: string;
begin
  try
    if (FIsEcritureFictive) then
      Result := FOmEcriture.EC_Piece
    else
      Result := FOmJournal.NextEC_Piece[FPeriode];
  except
    Result := '1';
  end;

end;

function TJournalVerrou.getPeriode: string;
begin
  result := FormatDateTime('yyyymm', FPeriode);
end;

/// <summary>
/// Verrouille la période du journal.
/// </summary>
procedure TJournalVerrou.Verrouiller;
var
  OmCptaApplication: IBSCptaApplication3;
  OmEcrituresJournal: IBICollection;
begin
  try
    OmCptaApplication := FOmJournal.Stream as IBSCptaApplication3;
    OmEcrituresJournal := OmCptaApplication.FactoryEcriture.QueryJournalPeriode(FOmJournal, FPeriode);

    if ((OmEcrituresJournal <> nil) and (OmEcrituresJournal.Count > 0)) then begin
      FOmEcriture := OmEcrituresJournal.Item[1] as IBOEcriture3;
      FOmEcriture.CouldModified;
      FIsEcritureFictive := False;
    end else begin
      FOmEcriture := OmCptaApplication.FactoryEcriture.Create as IBOEcriture3;
      FOmEcriture.Journal := FOmJournal;
      FOmEcriture.CompteG := FOmCompteGeneral;
      FOmEcriture.Date := FPeriode;
      FOmEcriture.EC_Intitule := 'Verrou';
      FOmEcriture.EC_Montant := 0.01;
      FOmEcriture.EC_Sens := EcritureSensTypeDebit;
      FOmEcriture.EC_Piece := FOmJournal.NextEC_Piece[FPeriode];
      FOMEcriture.Write();
      FOmEcriture.CouldModified;
      FIsEcritureFictive := True;
    end;

    FIsJournalVerrouille := True;
  except
    on E: Exception do Raise Exception.CreateFmt('Erreur en verrouillage du journal %s de %s : %s',
        [FOmJournal.JO_Num, FormatDateTime('mmmm yyyy', FPeriode), E.Message]);
  end;
end;

end.
