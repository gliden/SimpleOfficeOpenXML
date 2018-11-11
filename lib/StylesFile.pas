unit StylesFile;

interface

uses
  JvSimpleXml, CellFormat, Workbook, Sheets;

type
  TXlsxStylesFile = class(TObject)
  private
    FInternalFormatList: TXlsxDistinctFormatList;
  public
    constructor Create();
    destructor Destroy;override;

    procedure BuildFormatList(workbook: TXlsxWorkbook);
    procedure SaveToXml(basepath: String);
  end;

implementation

uses
  System.IOUtils, System.SysUtils, JclStreams;

{ TXlsxStylesFile }

procedure TXlsxStylesFile.BuildFormatList(workbook: TXlsxWorkbook);
var
  worksheet: TXlsxSheet;
begin
  FInternalFormatList.Clear;

  for worksheet in workbook.Sheets do
  begin
    worksheet.BuildFormatList(FInternalFormatList);
  end;
end;

constructor TXlsxStylesFile.Create;
begin
  FInternalFormatList := TXlsxDistinctFormatList.Create;
end;

destructor TXlsxStylesFile.Destroy;
begin
  FInternalFormatList.Free;
  inherited;
end;

procedure TXlsxStylesFile.SaveToXml(basepath: String);
var
  filename: string;
  xmlFile: TJvSimpleXML;
  styleSheetNode: TJvSimpleXMLElem;
  fontsNode: TJvSimpleXMLElem;
  fillsNode: TJvSimpleXMLElem;
  fillNode: TJvSimpleXMLElem;
  bordersNode: TJvSimpleXMLElem;
  borderNode: TJvSimpleXMLElem;
  cellStyleXfsNodes: TJvSimpleXMLElem;
  XfNode: TJvSimpleXMLElem;
  cellXfsNodes: TJvSimpleXMLElem;

  format: TXlsxCellFormat;
begin
  basePath := TPath.Combine(basepath, 'xl');
  ForceDirectories(basePath);
  filename := TPath.Combine(basepath, 'styles.xml');

  xmlFile := TJvSimpleXML.Create(nil);
  styleSheetNode := xmlFile.Root.Items.Add('styleSheet');
  styleSheetNode.Properties.Add('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');

  //Fonts
  fontsNode := styleSheetNode.Items.Add('fonts');
  fontsNode.Properties.Add('count', FInternalFormatList.Count);

  for format in FInternalFormatList do
  begin
    format.SaveToFontNode(fontsNode);
  end;

  //Fills
  fillsNode := styleSheetNode.Items.Add('fills');
  fillsNode.Properties.Add('count', 1);
  fillNode := fillsNode.Items.Add('fill');
  fillNode.Items.Add('patternFill').Properties.Add('patternType', 'none');

//  Borders
  bordersNode := styleSheetNode.Items.Add('borders');
  bordersNode.Properties.Add('count', FInternalFormatList.Count);
  for format in FInternalFormatList do
  begin
    format.SaveToBorderNode(bordersNode);
  end;

  //cellStylesXfs
  cellStyleXfsNodes := styleSheetNode.Items.Add('cellStyleXfs');
  cellStyleXfsNodes.Properties.Add('count', 1);
  XfNode := cellStyleXfsNodes.Items.Add('xf');
  XfNode.Properties.Add('numFmtId', 0);
  XfNode.Properties.Add('fontId', 0);
  XfNode.Properties.Add('fillId', 0);
  XfNode.Properties.Add('borderId', 0);

  //cellXfs
  cellXfsNodes := styleSheetNode.Items.Add('cellXfs');
  cellXfsNodes.Properties.Add('count', FInternalFormatList.Count);
  for format in FInternalFormatList do
  begin
    XfNode := cellXfsNodes.Items.Add('xf');
    XfNode.Properties.Add('numFmtId', 0);
    XfNode.Properties.Add('fontId', format.FormatId);
    XfNode.Properties.Add('fillId', 0);
    XfNode.Properties.Add('borderId', format.FormatId);
    XfNode.Properties.Add('xfId', 0);
    XfNode.Properties.Add('applyFont', 1);
  end;

  xmlFile.SaveToFile(filename, TJclStringEncoding.seUTF8);
  xmlFile.Free;
end;

(*
procedure TXlsxStylesFile.SaveToXml(basepath: String);
var
  filename: string;
  xmlFile: TJvSimpleXML;
  styleSheetNode: TJvSimpleXMLElem;
  fontsNode: TJvSimpleXMLElem;
  fontNode: TJvSimpleXMLElem;
  fillsNode: TJvSimpleXMLElem;
  fillNode: TJvSimpleXMLElem;
  bordersNode: TJvSimpleXMLElem;
  borderNode: TJvSimpleXMLElem;
  cellStyleXfsNodes: TJvSimpleXMLElem;
  XfNode: TJvSimpleXMLElem;
  cellXfsNodes: TJvSimpleXMLElem;
begin
  basePath := TPath.Combine(basepath, 'xl');
  ForceDirectories(basePath);
  filename := TPath.Combine(basepath, 'styles.xml');

  xmlFile := TJvSimpleXML.Create(nil);
  styleSheetNode := xmlFile.Root.Items.Add('styleSheet');
  styleSheetNode.Properties.Add('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');

  //Fonts
  fontsNode := styleSheetNode.Items.Add('fonts');
  fontsNode.Properties.Add('count', 2);

  fontNode := fontsNode.Items.Add('font');
  fontNode.Items.Add('sz').Properties.Add('val', '11');
  fontNode.Items.Add('color').Properties.Add('theme', '1');
  fontNode.Items.Add('name').Properties.Add('val', 'Calibri');
  fontNode.Items.Add('family').Properties.Add('val', '2');
  fontNode.Items.Add('scheme').Properties.Add('val', 'minor');

  fontNode := fontsNode.Items.Add('font');
  fontNode.Items.Add('b');
  fontNode.Items.Add('sz').Properties.Add('val', '11');
  fontNode.Items.Add('color').Properties.Add('theme', '1');
  fontNode.Items.Add('name').Properties.Add('val', 'Calibri');
  fontNode.Items.Add('family').Properties.Add('val', '2');
  fontNode.Items.Add('scheme').Properties.Add('val', 'minor');

  //Fills
  fillsNode := styleSheetNode.Items.Add('fills');
  fillsNode.Properties.Add('count', 1);
  fillNode := fillsNode.Items.Add('fill');
  fillNode.Items.Add('patternFill').Properties.Add('patternType', 'none');

  //Borders
  bordersNode := styleSheetNode.Items.Add('borders');
  bordersNode.Properties.Add('count', 2);

  borderNode := bordersNode.Items.Add('border');
  borderNode.Items.Add('left');
  borderNode.Items.Add('right');
  borderNode.Items.Add('top');
  borderNode.Items.Add('bottom');
  borderNode.Items.Add('diagonal');

  borderNode := bordersNode.Items.Add('border');
  borderNode.Items.Add('left');
  borderNode.Items.Add('right');
  borderNode.Items.Add('top');
  borderNode.Items.Add('bottom').Properties.Add('style', 'thick');
  borderNode.Items.Add('diagonal');

  //cellStylesXfs
  cellStyleXfsNodes := styleSheetNode.Items.Add('cellStyleXfs');
  cellStyleXfsNodes.Properties.Add('count', 1);
  XfNode := cellStyleXfsNodes.Items.Add('xf');
  XfNode.Properties.Add('numFmtId', 0);
  XfNode.Properties.Add('fontId', 0);
  XfNode.Properties.Add('fillId', 0);
  XfNode.Properties.Add('borderId', 0);

  //cellXfs
  cellXfsNodes := styleSheetNode.Items.Add('cellXfs');
  cellXfsNodes.Properties.Add('count', 2);
  XfNode := cellXfsNodes.Items.Add('xf');
  XfNode.Properties.Add('numFmtId', 0);
  XfNode.Properties.Add('fontId', 0);
  XfNode.Properties.Add('fillId', 0);
  XfNode.Properties.Add('borderId', 0);
  XfNode.Properties.Add('xfId', 0);

  XfNode := cellXfsNodes.Items.Add('xf');
  XfNode.Properties.Add('numFmtId', 0);
  XfNode.Properties.Add('fontId', 1);
  XfNode.Properties.Add('fillId', 0);
  XfNode.Properties.Add('borderId', 1);
  XfNode.Properties.Add('xfId', 0);
  XfNode.Properties.Add('applyFont', 1);


  xmlFile.SaveToFile(filename, TJclStringEncoding.seUTF8);
  xmlFile.Free;
end;*)

end.
