unit Workbook;

interface

uses
  Sheets, System.Generics.Collections;

type
  TXlsxWorkbook = class(TObject)
  private
    FSheets: TObjectList<TXlsxSheet>;
    procedure InternalSaveRelations(basePath: String);
  public
    constructor Create;
    destructor Destroy;override;

    property Sheets: TObjectList<TXlsxSheet> read FSheets;
    procedure SaveToXml(basepath: String);
    procedure AddSheet(name: String);
  end;

implementation

uses
  JvSimpleXml, System.IOUtils, System.SysUtils;

{ TXlsxWorkbook }

procedure TXlsxWorkbook.AddSheet(name: String);
var
  sheet: TXlsxSheet;
begin
  sheet := TXlsxSheet.Create;
  sheet.Name := name;
  sheet.Id := FSheets.Count + 1;
  FSheets.Add(sheet);
end;

constructor TXlsxWorkbook.Create;
begin
  FSheets := TObjectList<TXlsxSheet>.Create;
end;

destructor TXlsxWorkbook.Destroy;
begin
  FSheets.Free;
  inherited;
end;

procedure TXlsxWorkbook.InternalSaveRelations(basePath: String);
var
  xmlExport: TJvSimpleXML;
  relationshipsNode: TJvSimpleXMLElem;
  relationShipNode: TJvSimpleXMLElem;
  sheet: TXlsxSheet;
  filename: string;
begin
  basePath := TPath.Combine(basePath, '_rels');
  ForceDirectories(basePath);
  filename := TPath.Combine(basePath, 'workbook.xml.rels');

  xmlExport := TJvSimpleXML.Create(nil);
  relationshipsNode := xmlExport.Root.Items.Add('Relationships');
  relationshipsNode.Properties.Add('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

  for sheet in FSheets do
  begin
    relationShipNode := relationshipsNode.items.Add('Relationship');
    relationShipNode.Properties.Add('Id', sheet.Reference);
    relationShipNode.Properties.Add('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet');
    relationShipNode.Properties.Add('Target', Format('worksheets/sheet%d.xml', [sheet.Id]));
  end;

  xmlExport.SaveToFile(filename);
  xmlExport.Free;
end;

procedure TXlsxWorkbook.SaveToXml(basepath: String);
var
  xmlExport: TJvSimpleXML;
  workbookNode: TJvSimpleXMLElem;
  sheetsNode: TJvSimpleXMLElem;
  sheet: TXlsxSheet;
  filename: string;
begin
  basePath := TPath.Combine(basepath, 'xl');
  ForceDirectories(basePath);
  filename := TPath.Combine(basepath, 'workbook.xml');

  xmlExport := TJvSimpleXML.Create(nil);
  workbookNode := xmlExport.Root.Items.Add('workbook');
  workbookNode.Properties.Add('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
  workbookNode.Properties.Add('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');

  workbookNode.Items.Add('bookViews').Items.Add('workbookView').Properties.Add('activeTab', 1);

  sheetsNode := workbookNode.Items.Add('sheets');

  for sheet in FSheets do
  begin
    sheet.SaveToWorkbookXmlNode(sheetsNode);
  end;

  xmlExport.SaveToFile(filename);
  xmlExport.Free;

  InternalSaveRelations(basepath);

  basePath := TPath.Combine(basepath, 'worksheets');
  ForceDirectories(basePath);
  for sheet in FSheets do
  begin
    sheet.SaveToXml(basepath);
  end;
end;

end.
