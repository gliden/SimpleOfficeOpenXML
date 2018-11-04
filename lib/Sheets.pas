unit Sheets;

interface

uses
  JvSimpleXml, Cell, System.Generics.Collections;

type
  TXlsxSheet = class(TObject)
  private
    FName: String;
    FId: Integer;
    FCells: TObjectList<TXlsxCell>;
    function GenerateReference: String;
    function getCell(col, row: Integer): TXlsxCell;
  public
    constructor Create;
    destructor Destroy;override;

    property Id: Integer read FId write FId;
    property Reference: String read GenerateReference;
    property Name: String read FName write FName;
    property Cell[col, row: Integer]: TXlsxCell read getCell;

    procedure SaveToWorkbookXmlNode(node: TJvSimpleXMLElem);
    procedure SaveToXml(basepath: String);
  end;

implementation

uses
  System.SysUtils, System.IOUtils;

{ TXlsxSheet }

constructor TXlsxSheet.Create;
begin
  FCells := TObjectList<TXlsxCell>.Create;
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
    Result := TXlsxCell.Create(row, col);
    FCells.Add(Result);
  end;
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
  for tmpCell in FCells do
  begin
    tmpCell.SaveToWorksheetXmlNode(sheetDataNode);
  end;

  xmlExport.SaveToFile(filename);
  xmlExport.Free;
end;

end.
