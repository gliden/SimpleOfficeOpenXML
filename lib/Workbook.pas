unit Workbook;

interface

uses
  Sheets, System.Generics.Collections, CellFormat, SharedStrings;

type
  TXlsxWorkbook = class(TObject)
  private
    FSheets: TObjectList<TXlsxSheet>;
    FDefaultFormat: TXlsxCellFormat;
    FSharedStrings: TXlsxSharedStrings;
    procedure InternalSaveRelations(basePath: String);
  public
    constructor Create;
    destructor Destroy;override;

    property DefaultFormat: TXlsxCellFormat read FDefaultFormat;
    property Sheets: TObjectList<TXlsxSheet> read FSheets;
    procedure SaveToXml(basepath: String);
    procedure LoadFromXml(basepath: String);
    procedure AddSheet(name: String);
  end;

implementation

uses
  JvSimpleXml, System.IOUtils, System.SysUtils, JclStreams, Helper, StylesFile;

{ TXlsxWorkbook }

procedure TXlsxWorkbook.AddSheet(name: String);
var
  sheet: TXlsxSheet;
begin
  sheet := TXlsxSheet.Create(FDefaultFormat, FSharedStrings);
  sheet.Name := name;
  sheet.Id := FSheets.Count + 1;
  FSheets.Add(sheet);
end;

constructor TXlsxWorkbook.Create;
begin
  FSheets := TObjectList<TXlsxSheet>.Create;
  FDefaultFormat := TXlsxCellFormat.Create;
  FSharedStrings := TXlsxSharedStrings.Create;
end;

destructor TXlsxWorkbook.Destroy;
begin
  FDefaultFormat.Free;
  FSheets.Free;
  FSharedStrings.Free;
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

  relationShipNode := relationshipsNode.items.Add('Relationship');
  relationShipNode.Properties.Add('Id', Format('rId%d', [FSheets.Count+1]));
  relationShipNode.Properties.Add('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles');
  relationShipNode.Properties.Add('Target', 'styles.xml');

  relationShipNode := relationshipsNode.items.Add('Relationship');
  relationShipNode.Properties.Add('Id', Format('rId%d', [FSheets.Count+2]));
  relationShipNode.Properties.Add('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings');
  relationShipNode.Properties.Add('Target', 'sharedStrings.xml');

  for sheet in FSheets do
  begin
    relationShipNode := relationshipsNode.items.Add('Relationship');
    relationShipNode.Properties.Add('Id', sheet.Reference);
    relationShipNode.Properties.Add('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet');
    relationShipNode.Properties.Add('Target', Format('worksheets/sheet%d.xml', [sheet.Id]));
  end;

  xmlExport.SaveToFile(filename, TJclStringEncoding.seUTF8);
  xmlExport.Free;
end;

procedure TXlsxWorkbook.LoadFromXml(basepath: String);
var
  xmlImport: TJvSimpleXML;
  sheetsNode: TJvSimpleXMLElem;
  sheet: TXlsxSheet;
  filename: string;
  i: Integer;
  sheetNode: TJvSimpleXMLElem;
  fStyles: TXlsxStylesFile;

begin
  basePath := TPath.Combine(basepath, 'xl');
  filename := TPath.Combine(basepath, 'workbook.xml');

  FSharedStrings.LoadFromXml(basepath);

  fStyles := TXlsxStylesFile.Create;
  fStyles.LoadFromXML(basepath);

  xmlImport := TJvSimpleXML.Create(nil);
  xmlImport.LoadFromFile(filename, TJclStringEncoding.seUTF8);
  sheetsNode := xmlImport.Root.Items.ItemNamed['sheets'];
  for i := 0 to sheetsNode.Items.Count-1 do
  begin
    sheetNode := sheetsNode.Items[i];
    { TODO : Load worksheet references from the workbook.xml.rels }
    if SameText(sheetNode.Name, 'sheet') then
    begin
      sheet := TXlsxSheet.Create(FDefaultFormat, FSharedStrings);
      sheet.LoadFromWorkbookXmlNode(sheetNode);
      sheet.Id := i+1; //Just a little workaround
      sheet.LoadFromXml(TPath.Combine(basepath, 'worksheets'));
      FSheets.Add(sheet);
    end;
  end;
  fStyles.SetFormatList(Self);
  fStyles.Free;

  xmlImport.Free;
end;

procedure TXlsxWorkbook.SaveToXml(basepath: String);
var
  xmlExport: TJvSimpleXML;
  workbookNode: TJvSimpleXMLElem;
  sheetsNode: TJvSimpleXMLElem;
  sheet: TXlsxSheet;
  filename: string;
  worksheetPath: string;
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

  xmlExport.SaveToFile(filename, TJclStringEncoding.seUTF8);
  xmlExport.Free;

  InternalSaveRelations(basepath);

  worksheetPath := TPath.Combine(basepath, 'worksheets');
  ForceDirectories(worksheetPath);
  for sheet in FSheets do
  begin
    sheet.SaveToXml(worksheetPath);
  end;

  FSharedStrings.SaveToXml(basepath);
end;

end.

