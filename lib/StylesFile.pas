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
    procedure SetFormatList(workbook: TXlsxWorkbook);
    procedure SaveToXml(basepath: String);
    procedure LoadFromXML(basepath: String);
  end;

implementation

uses
  System.IOUtils, System.SysUtils, JclStreams, Helper;

{ TXlsxStylesFile }

procedure TXlsxStylesFile.BuildFormatList(workbook: TXlsxWorkbook);
var
  worksheet: TXlsxSheet;
begin
  FInternalFormatList.Clear;
  FInternalFormatList.AddIfNotExists(workbook.DefaultFormat);

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

procedure TXlsxStylesFile.LoadFromXML(basepath: String);
var
  fFontList : TFontList;
  fBorderList : TBorderList;
  fFillList : TFillList;
  fFont : TXlsxCellFont;
  fBorder : TXlsxBorder;
  fFill : TXlsxCellFill;
  xmlImport: TJvSimpleXML;
  fontsNode : TJvSimpleXMLElem;
  bordersNode : TJvSimpleXMLElem;
  fillsNode : TJvSimpleXMLElem;
  i         : integer;
  filename : String;
  cellXfsNode: TJvSimpleXMLElem;
  cellFormat: TXlsxCellFormat;
  xfNode: TJvSimpleXMLElem;
  fontId: Int64;
  fillId: Int64;

  procedure AddFont(fontNode: TJvSimpleXMLElem);
  var
    indexValue : Integer;
  begin
    fFont := TXlsxCellFont.Create;
    fFontList.Add(fFont);
    if assigned(fontNode.Items.ItemNamed['sz']) then
    begin
      fFont.Size := fontNode.Items.ItemNamed['sz'].Properties.itemNamed['val'].intValue;
    end;
    if assigned(fontNode.Items.ItemNamed['name']) then
    begin
      fFont.FontName := fontNode.Items.ItemNamed['name'].Properties.ItemNamed['val'].Value;
    end;
    if assigned(fontNode.Items.ItemNamed['color']) then
    begin
      if Assigned(fontNode.Items.ItemNamed['color'].Properties.ItemNamed['theme']) then
      begin
        fFont.ColorTheme := fontNode.Items.ItemNamed['color'].Properties.ItemNamed['theme'].intValue;
      end
      else if Assigned(fontNode.Items.ItemNamed['color'].Properties.ItemNamed['indexed']) then
      begin
        indexValue := fontNode.Items.ItemNamed['color'].Properties.ItemNamed['indexed'].intValue;
        fFont.SetIndexedColor(indexValue);
      end;
    end;
    if assigned(fontNode.Items.ItemNamed['family']) then
    begin
      fFont.Family := fontNode.Items.ItemNamed['family'].Properties.ItemNamed['val'].intValue;
    end;
    if assigned(fontNode.items.ItemNamed['b']) then fFont.Style := fFont.Style + [xfsBold];
    if assigned(fontNode.items.ItemNamed['i']) then fFont.Style := fFont.Style + [xfsItalic];
    if assigned(fontNode.items.ItemNamed['u']) then fFont.Style := fFont.Style + [xfsUnderline];
  end;

  procedure AddBorder(boderNode: TJvSimpleXMLElem);
  begin
    fBorder := TXlsxBorder.Create;
    fBorderList.Add(fBorder);
    { TODO : Load borders }
  end;

  procedure AddFill(fillNode : TJvSimpleXMLElem);
  var
    rgbValue : String;
    indexValue : Integer;
    doubleValue : double;
    patternNode : TJvSimpleXMLElem;
    hStr        : String;
  begin
    fFill := TXlsxCellFill.Create;
    fFillList.Add(fFill);
    patternNode := fillNode.Items.ItemNamed['patternFill'];
    if assigned(patternNode) then
    begin
      fFill.SetPatternTypeByName(patternNode.Properties.ItemNamed['patternType'].Value);

      if assigned(patternNode.Items.ItemNamed['fgColor']) then
      begin
        if assigned(patternNode.Items.ItemNamed['fgColor'].Properties.ItemNamed['rgb']) then
        begin
          rgbValue := patternNode.Items.ItemNamed['fgColor'].Properties.ItemNamed['rgb'].Value;
          fFill.SetfgColorByRGB(rgbValue);
        end
        else if assigned(patternNode.Items.ItemNamed['fgColor'].Properties.ItemNamed['indexed']) then
        begin
          indexValue := patternNode.Items.ItemNamed['fgColor'].Properties.ItemNamed['indexed'].IntValue;
          fFill.SetFgIndexedColor(indexValue);
        end;
        if assigned(patternNode.Items.ItemNamed['fgColor'].Properties.ItemNamed['theme']) then
        begin
          fFill.ThemeFgColor := patternNode.Items.ItemNamed['fgColor'].Properties.ItemNamed['theme'].IntValue;
        end;
        if assigned(patternNode.Items.ItemNamed['fgColor'].Properties.ItemNamed['tint']) then
        begin
          hStr := patternNode.Items.ItemNamed['fgColor'].Properties.ItemNamed['tint'].Value;
          hStr := StringReplace(hStr, '.', ',', []);
          if TryStrToFloat(hStr, doubleValue) then fFill.TintFgColor := doubleValue;
        end;
      end;
      if assigned(patternNode.Items.ItemNamed['bgColor']) then
      begin
        if assigned(patternNode.Items.ItemNamed['bgColor'].Properties.ItemNamed['rgb']) then
        begin
          rgbValue := patternNode.Items.ItemNamed['bgColor'].Properties.ItemNamed['rgb'].Value;
          fFill.SetbgColorByRGB(rgbValue);
        end
        else if assigned(patternNode.Items.ItemNamed['bgColor'].Properties.ItemNamed['indexed']) then
        begin
          indexValue := patternNode.Items.ItemNamed['bgColor'].Properties.ItemNamed['indexed'].IntValue;
          fFill.SetBgIndexedColor(indexValue);
        end;
        if assigned(patternNode.Items.ItemNamed['bgColor'].Properties.ItemNamed['theme']) then
        begin
          fFill.ThemeBgColor := patternNode.Items.ItemNamed['bgColor'].Properties.ItemNamed['theme'].IntValue;
        end;
        if assigned(patternNode.Items.ItemNamed['bgColor'].Properties.ItemNamed['tint']) then
        begin
          hStr := patternNode.Items.ItemNamed['bgColor'].Properties.ItemNamed['tint'].Value;
          hStr := StringReplace(hStr, '.', ',', []);
          if TryStrToFloat(hStr, doubleValue) then fFill.TintBgColor := doubleValue;
        end;
      end;
    end;
  end;

begin
  FInternalFormatList.Clear;

  fFontList := TFontList.Create;
  fBorderList := TBorderList.Create;
  fFillList := TFillList.Create;

  xmlImport := TJvSimpleXML.Create(nil);
  filename := TPath.Combine(basepath, 'styles.xml');
  xmlImport.LoadFromFile(filename, TJclStringEncoding.seUTF8);
  fontsNode := xmlImport.Root.Items.ItemNamed['fonts'];
  if Assigned(fontsNode) then
  begin
    for i := 0 to fontsNode.Items.Count-1 do
    begin
      AddFont(fontsNode.Items[i]);
    end;
  end;
  bordersNode := xmlImport.Root.Items.ItemNamed['borders'];
  if Assigned(bordersNode) then
  begin
    for i := 0 to bordersNode.Items.Count-1 do
    begin
      AddBorder(bordersNode.Items[i]);
    end;
  end;
  fillsNode := xmlImport.Root.Items.ItemNamed['fills'];
  if Assigned(fillsNode) then
  begin
    for i := 0 to fillsNode.Items.Count-1 do
    begin
      AddFill(fillsNode.Items[i]);
    end;
  end;

  // ToDo: Formate zusammensetzen
  cellXfsNode := xmlImport.Root.Items.ItemNamed['cellXfs'];
  for i := 0 to cellXfsNode.Items.Count-1 do
  begin
    xfNode := cellXfsNode.Items[i];
    fontId := xfNode.Properties.ItemNamed['fontId'].IntValue;
    fillId := xfNode.Properties.ItemNamed['fillId'].IntValue;

    cellFormat := TXlsxCellFormat.Create;
    cellFormat.FormatId := i;
    cellFormat.Font.Assign(fFontList[fontId]);
    cellFormat.Fill.Assign(fFillList[fillId]);
    FInternalFormatList.Add(cellFormat);
  end;


  fFillList.Free;
  fBorderList.Free;
  fFontList.Free;
end;

procedure TXlsxStylesFile.SaveToXml(basepath: String);
var
  filename: string;
  xmlFile: TJvSimpleXML;
  styleSheetNode: TJvSimpleXMLElem;
  fontsNode: TJvSimpleXMLElem;
  fillsNode: TJvSimpleXMLElem;
  bordersNode: TJvSimpleXMLElem;
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

  //Fills
  fillsNode := styleSheetNode.Items.Add('fills');
  fillsNode.Properties.Add('count', FInternalFormatList.Count);

  for format in FInternalFormatList do
  begin
    format.SaveToFontNode(fontsNode);
    format.SaveToFillsNode(fillsNode);
  end;

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
    XfNode.Properties.Add('fillId', format.FormatId);
    XfNode.Properties.Add('borderId', format.FormatId);
    XfNode.Properties.Add('xfId', 0);
    XfNode.Properties.Add('applyFont', 1);
  end;

  xmlFile.SaveToFile(filename, TJclStringEncoding.seUTF8);
  xmlFile.Free;
end;

procedure TXlsxStylesFile.SetFormatList(workbook: TXlsxWorkbook);
var
  worksheet: TXlsxSheet;
begin
  workbook.DefaultFormat.Assign(FInternalFormatList[0]);
  for worksheet in workbook.Sheets do
  begin
    worksheet.SetFormatList(FInternalFormatList);
  end;
end;

end.
