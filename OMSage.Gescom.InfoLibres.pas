unit OMSage.Gescom.InfoLibres;

{*

Classe delphi depuis le code cs deThierry JOUVE - Sage Service Consulting

*}

interface

uses
  System.SysUtils, System.Classes, System.Generics.Collections,
  {$INCLUDE OmSageLib.inc};

type

  TInfoLibreLigneDoc = class
    private
      FDictionary: TDictionary<Integer, OleVariant>;
      FIPMDocument: IPMDocument;
      FIntituleInfoLibre: string;
      function GetValue(const Idx: Integer; var Value: OleVariant): boolean;
    public
      Constructor Create(AIPMDocument: IPMDocument; AIntituleInfoLibre: string); reintroduce;
      Destructor Destroy; override;
      procedure SetValue(const Value: OleVariant; BeforeLineCreated: boolean = false);
      procedure ValidateValues;
    end;

implementation

uses
  System.Variants;

{ TInfoLibreLigneDoc }

constructor TInfoLibreLigneDoc.Create(AIPMDocument: IPMDocument;
  AIntituleInfoLibre: string);
begin
  FDictionary := TDictionary<Integer, OleVariant>.Create;
  FIPMDocument := AIPMDocument;
  FIntituleInfoLibre := AIntituleInfoLibre;
end;

destructor TInfoLibreLigneDoc.Destroy;
begin
  FDictionary.Free;
  inherited;
end;

function TInfoLibreLigneDoc.GetValue(const Idx: Integer; var Value: OleVariant): boolean;
begin
  Result := FDictionary.TryGetValue(Idx, Value);
end;

procedure TInfoLibreLigneDoc.SetValue(const Value: OleVariant; BeforeLineCreated: boolean = false);
var
  Idx: integer;
begin
  if (Value <> Unassigned) then begin
    Idx := FIPMDocument.Document.FactoryDocumentLigne.List.Count;
    if (BeforeLineCreated) then Inc(Idx);
    FDictionary.AddOrSetValue(Idx, Value);
  end;
end;

procedure TInfoLibreLigneDoc.ValidateValues;
var
  Idx: integer;
  Value: OleVariant;
  Ligne: IBODocumentLigne3;
begin
  if (FIPMDocument.DocumentResult <> nil) then begin
    for Idx := 1 to FIPMDocument.DocumentResult.FactoryDocumentLigne.List.Count do begin
      if GetValue(Idx, Value) then begin
        Ligne := FIPMDocument.DocumentResult.FactoryDocumentLigne.List.Item[Idx] as IBODocumentLigne3;
        Ligne.InfoLibre[FIntituleInfoLibre] := Value;
        Ligne.Write;
      end;
    end;
  end;
end;

end.
