unit SharedStrings;

interface

uses
  System.Generics.Collections;

type
  TXlsxString = class(TObject)
  private
    FId: Integer;
    FValue: String;
  public
    property Id: Integer read FId;
    property Value: String read FValue;
  end;

  TXlsxSharedStrings = class(TObject)
  private
    FStrings: TObjectList<TXlsxString>;
  public
    constructor Create;
    destructor Destroy;override;

    procedure LoadFromXml(basePath: String);
    procedure SaveToXml(basePath: String);

    function AddString(value: String): Integer;
    function ValueById(id: Integer): String;
  end;

implementation

uses
  System.SysUtils, JvSimpleXml, System.IOUtils;

{ TXlsxSharedStrings }

function TXlsxSharedStrings.AddString(value: String): Integer;
var
  s: TXlsxString;
begin
  Result := -1;
  for s in FStrings do
  begin
    if s.Value = value then
    begin
      Result := s.Id;
      break;
    end;
  end;

  if Result = -1 then
  begin
    s := TXlsxString.Create;
    s.FId := FStrings.Count;
    s.FValue := value;
    FStrings.Add(s);
    Result := s.Id;
  end;
end;

constructor TXlsxSharedStrings.Create;
begin
  FStrings := TObjectList<TXlsxString>.Create;
end;

destructor TXlsxSharedStrings.Destroy;
begin
  FStrings.Free;
  inherited;
end;

procedure TXlsxSharedStrings.LoadFromXml(basePath: String);
var
  xmlImport: TJvSimpleXML;
  i: Integer;
  siNode: TJvSimpleXMLElem;
  sharedString: TXlsxString;
begin
  xmlImport := TJvSimpleXML.Create(nil);
  xmlImport.LoadFromFile(TPath.Combine(basePath, 'sharedStrings.xml'));
  FStrings.Clear;
  for i := 0 to xmlImport.Root.Items.Count-1 do
  begin
    siNode := xmlImport.Root.Items[i];
    sharedString := TXlsxString.Create;
    sharedString.FId := FStrings.Count;
    sharedString.FValue := siNode.Items.ItemNamed['t'].Value;
    FStrings.Add(sharedString);
  end;
  xmlImport.Free;
end;

procedure TXlsxSharedStrings.SaveToXml(basePath: String);
var
  xmlExport: TJvSimpleXML;
  filename: string;
  sstNode: TJvSimpleXMLElem;
  s: TXlsxString;
begin
  filename := TPath.Combine(basePath, 'sharedStrings.xml');

  xmlExport := TJvSimpleXML.Create(nil);
  sstNode := xmlExport.Root.Items.Add('sst');
  sstNode.Properties.Add('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
  sstNode.Properties.Add('count', FStrings.Count);
  sstNode.Properties.Add('uniqueCount', FStrings.Count);

  for s in FStrings do
  begin
    sstNode.Items.Add('si').Items.Add('t', s.Value);
  end;

  xmlExport.SaveToFile(filename);
  xmlExport.Free;
end;

function TXlsxSharedStrings.ValueById(id: Integer): String;
begin
  Result := '';
  if FStrings.Count>id then
  begin
    Result := FStrings[id].Value;
  end;
end;

end.
