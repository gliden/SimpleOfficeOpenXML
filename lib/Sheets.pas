unit Sheets;

interface

uses
  JvSimpleXml, Cell, System.Generics.Collections, CellFormat, SharedStrings;

type
  TXlsxSheet = class(TObject)
  private
    FName: String;
    FId: Integer;
    FCells: TObjectList<TXlsxCell>;
    FDefaultFormat: TXlsxCellFormat;
    FSharedStrings: TXlsxSharedStrings;
    function GenerateReference: String;
    function getCell(col, row: Integer): TXlsxCell;
  public
    constructor Create(defaultFormat: TXlsxCellFormat; sharedStrings: TXlsxSharedStrings);
    destructor Destroy;override;
    procedure BuildFormatList(list: TXlsxDistinctFormatList);

    property Id: Integer read FId write FId;
    property Reference: String read GenerateReference;
    property Name: String read FName write FName;
    property Cell[col, row: Integer]: TXlsxCell read getCell;

    procedure SaveToWorkbookXmlNode(node: TJvSimpleXMLElem);
    procedure LoadFromWorkbookXmlNode(node: TJvSimpleXMLElem);
    procedure SaveToXml(basepath: String);
    procedure LoadFromXml(basepath: String);
  end;

implementation

uses
  System.SysUtils, System.IOUtils, JclStreams, System.Generics.Defaults,
  System.Math, Helper;

{ TXlsxSheet }

procedure TXlsxSheet.BuildFormatList(list: TXlsxDistinctFormatList);
var
  tmpCell: TXlsxCell;
begin
  for tmpCell in FCells do
  begin
    tmpCell.Format.FormatId := list.AddIfNotExists(tmpCell.Format);
  end;
end;

constructor TXlsxSheet.Create(defaultFormat: TXlsxCellFormat; sharedStrings: TXlsxSharedStrings);
begin
  FCells := TObjectList<TXlsxCell>.Create;
  FDefaultFormat := defaultFormat;
  FSharedStrings := sharedStrings;
end;

destructor TXlsxSheet.Destroy;
begin
  FCells.Free;
  inherited;
end;

function TXlsxSheet.GenerateReference: String;
begin
  Result := Format('rId%d', [fId]);
end;

function TXlsxSheet.getCell(col, row: Integer): TXlsxCell;
var
  tmpCell: TXlsxCell;
begin
  Result := nil;
  for tmpCell in FCells do
  begin
    if (tmpCell.Row = row) and (tmpCell.Col = col) then
    begin
      Result := tmpCell;
      break;
    end;
  end;

  if Result = nil then
  begin
    Result := TXlsxCell.Create(row, col, FDefaultFormat, FSharedStrings);
    FCells.Add(Result);
  end;
end;

procedure TXlsxSheet.LoadFromWorkbookXmlNode(node: TJvSimpleXMLElem);
begin
  FName := node.Properties.ItemNamed['name'].Value;
  FId := node.Properties.ItemNamed['sheetId'].IntValue;
end;

procedure TXlsxSheet.LoadFromXml(basepath: String);
var
  filename: string;
  xmlImport: TJvSimpleXML;
  sheetDataNode: TJvSimpleXMLElem;
  i: Integer;
  rowNode: TJvSimpleXMLElem;
  cellNode: TJvSimpleXMLElem;
  col: Integer;
  x: Integer;
  row: Integer;
  typeProperty: TJvSimpleXMLProp;
  formulaNode: TJvSimpleXMLElem;
begin
  filename := TPath.Combine(basepath, Format('sheet%d.xml', [FId]));

  xmlImport := TJvSimpleXML.Create(nil);
  xmlImport.LoadFromFile(filename);

  sheetDataNode := xmlImport.Root.Items.ItemNamed['sheetData'];
  for i := 0 to sheetDataNode.Items.Count-1 do
  begin
    rowNode := sheetDataNode.Items[i];
    for x := 0 to rowNode.Items.Count-1 do
    begin
      cellNode := rowNode.Items[x];
      col := ColReferenceToNumber(cellNode.Properties.ItemNamed['r'].Value);
      row := rowNode.Properties.ItemNamed['r'].IntValue;
      typeProperty := cellNode.Properties.ItemNamed['t'];
      formulaNode := cellNode.Items.ItemNamed['f'];
      
      if (typeProperty <> nil) then
      begin
        if SameText(typeProperty.Value, 'inlineStr') then
        begin
          Cell[col, row].Value := cellNode.Items.ItemNamed['is'].Items.ItemNamed['t'].Value;
        end else
        if SameText(typeProperty.Value, 's') then
        begin
          Cell[col, row].Value := FSharedStrings.ValueById(cellNode.Items.ItemNamed['v'].IntValue);
        end;
      end else
      if formulaNode <> nil then
      begin
        Cell[col, row].Formula := formulaNode.Value;
      end;
    end;
  end;

  xmlImport.Free;
end;

procedure TXlsxSheet.SaveToWorkbookXmlNode(node: TJvSimpleXMLElem);
var
  sheetNode: TJvSimpleXMLElem;
begin
  sheetNode := node.Items.Add('sheet');
  sheetNode.Properties.Add('name', FName);
  sheetNode.Properties.Add('sheetId', FId);
  sheetNode.Properties.Add('r:id', GenerateReference);
end;

procedure TXlsxSheet.SaveToXml(basepath: String);
var
  xmlExport: TJvSimpleXML;
  workSheetNode: TJvSimpleXMLElem;
  sheetDataNode: TJvSimpleXMLElem;
  filename: string;
  tmpCell: TXlsxCell;
begin
  filename := TPath.Combine(basepath, Format('sheet%d.xml', [FId]));

  xmlExport := TJvSimpleXML.Create(nil);
  workSheetNode := xmlExport.Root.Items.Add('worksheet');
  workSheetNode.Properties.Add('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
  workSheetNode.Properties.Add('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');

  sheetDataNode := workSheetNode.Items.Add('sheetData');
  FCells.Sort(
    TComparer<TXlsxCell>.Construct(
      function(const cell1, cell2: TXlsxCell): Integer
      begin
        Result := CompareValue(cell1.Row, cell2.Row);
        if Result = 0 then Result := CompareValue(cell1.Col, cell2.Col);
      end));
  for tmpCell in FCells do
  begin
    tmpCell.SaveToWorksheetXmlNode(sheetDataNode);
  end;

  xmlExport.SaveToFile(filename, TJclStringEncoding.seUTF8);
  xmlExport.Free;
end;

end.
