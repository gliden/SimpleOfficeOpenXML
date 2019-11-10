unit CellFormat;

interface

uses
  System.Generics.Collections, JvSimpleXml, Vcl.Graphics;

type
  TXlsxFontStyle = (xfsBold, xfsItalic, xfsUnderline);
  TXlsxFontStyles = set of TXlsxFontStyle;


  TXlsxBorderStyle = (xbsNone, xbsMediumDashDotDot, xbsHair, xbsSlantDashDot, xbsDotted, xbsMediumDashDot, xbsDashDotDot, xbsMediumDashed, xbsDashDot, xbsMedium, xbsDashed, xbsThick, xbsThin, xbsDouble);
  TXlsxPatternType = (xptNone, xptGray125, xptSolid);

  TXlsxBorder = class(TObject)
  private
    FStyle: TXlsxBorderStyle;
    function getBorderStyleToString: String;
  public
    constructor Create;
    procedure SaveToNode(nodeName: String; node: TJvSimpleXMLElem);
    procedure SetPatternTypeByName(StyleName: String);
    property Style: TXlsxBorderStyle read FStyle write FStyle;
  end;

  TXlsxCellFont = class(TObject)
  private
    FColorTheme: Integer;
    FSize: Integer;
    FFontName: String;
    FFamily: Integer;
    FStyle: TXlsxFontStyles;
    fIndexedColor : Integer;
  public
    constructor Create;
    procedure Assign(value: TXlsxCellFont);
    function IsSame(font: TXlsxCellFont): Boolean;
    procedure SetIndexedColor(idx: Integer);

    property Size: Integer read FSize write FSize;
    property ColorTheme: Integer read FColorTheme write FColorTheme;
    property FontName: String read FFontName write FFontName;
    property Family: Integer read FFamily write FFamily;
    property Style: TXlsxFontStyles read FStyle write FStyle;
  end;

  TXlsxCellFill = class(TObject)
  private
    FPatternType: TXlsxPatternType;
    fFgColor: TColor;
    fBgColor: TColor;
    fThemeFgColor : integer;
    fThemeBgColor : Integer;
    fTintFgColor : double;
    fTintBgColor : double;
    function fgToRGB: String;
    function bgToRGB: String;
  public
    constructor Create;
    procedure Assign(value: TXlsxCellFill);
    function IsSame(value: TXlsxCellFill): Boolean;
    procedure SetPatternTypeByName(PatternName: String);
    procedure SetFgIndexedColor(idx: integer);
    procedure SetBgIndexedColor(idx: integer);
    function SetbgColorByRGB(rgb: String): String;
    function SetfgColorByRGB(rgb: String): String;

    property PatternType: TXlsxPatternType read FPatternType write FPatternType;
    property fgColor: TColor read fFgColor write fFgColor;
    property bgColor: TColor read fBgColor write fBgColor;
    property ThemeFgColor : integer read fThemeFgColor write fThemeFgColor;
    property ThemeBgColor : Integer read fThemeBgColor write fThemeBgColor;
    property TintFgColor : double read fTintFgColor write fTintFgColor;
    property TintBgColor : double read fTintBgColor write fTintBgColor;
  end;

  TXlsxCellFormat = class(TObject)
  private
    FFont: TXlsxCellFont;
    FFormatId: Integer;
    FBottomBorder: TXlsxBorderStyle;
    FTopBorder: TXlsxBorderStyle;
    FDiagonalBorder: TXlsxBorderStyle;
    FLeftBorder: TXlsxBorderStyle;
    FRightBorder: TXlsxBorderStyle;
    FFill: TXlsxCellFill;
    function HasSameBorders(format: TXlsxCellFormat): Boolean;
    procedure SaveBorder(nodename: String; style: TXlsxBorderStyle; node: TJvSimpleXMLElem);
  public
    constructor Create;
    destructor Destroy; override;
    procedure Assign(value: TXlsxCellFormat);

    function IsSame(format: TXlsxCellFormat): Boolean;
    procedure SaveToFillsNode(node: TJvSimpleXMLElem);
    procedure SaveToFontNode(node: TJvSimpleXMLElem);
    procedure SaveToBorderNode(node: TJvSimpleXMLElem);

    property FormatId: Integer read FFormatId write FFormatId;
    property Font: TXlsxCellFont read FFont;
    property Fill: TXlsxCellFill read FFill;

    property LeftBorder: TXlsxBorderStyle read FLeftBorder write FLeftBorder;
    property RightBorder: TXlsxBorderStyle read FRightBorder write FRightBorder;
    property TopBorder: TXlsxBorderStyle read FTopBorder write FTopBorder;
    property BottomBorder: TXlsxBorderStyle read FBottomBorder write FBottomBorder;
    property DiagonalBorder: TXlsxBorderStyle read FDiagonalBorder write FDiagonalBorder;
  end;

  TXlsxDistinctFormatList = class(TList<TXlsxCellFormat>)
  public
    function AddIfNotExists(format: TXlsxCellFormat): Integer;
    function getFormatById(id: Integer): TXlsxCellFormat;
  end;

  TFontList = class(TList<TXlsxCellFont>);
  TBorderList = class(TList<TXlsxBorder>);
  TFillList = class(TList<TXlsxCellFill>);

implementation

uses
  ColorRecHelper, System.SysUtils, IndexedColors;

const XlsxBorderStyleName: array[TXlsxBorderStyle] of String = ('none', 'mediumDashDotDot', 'hair', 'slantDashDot',
                                                                'dotted', 'mediumDashDot', 'dashDotDot', 'mediumDashed',
                                                                'dashDot', 'medium', 'dashed', 'thick', 'thin', 'double');

const XlsxFillPatternTypeName: array[TXlsxPatternType] of String = ('none', 'gray125', 'solid');

{ TXlsxCellFormat }

procedure TXlsxCellFormat.Assign(value: TXlsxCelLFormat);
begin
  FFont.Assign(value.Font);
  FFill.Assign(value.Fill);
  FFormatId := value.FormatId;
  FBottomBorder := value.BottomBorder;
  FTopBorder := value.TopBorder;
  FDiagonalBorder := value.DiagonalBorder;
  FLeftBorder := value.LeftBorder;
  FRightBorder := value.RightBorder;
end;

constructor TXlsxCellFormat.Create;
begin
  FFont := TXlsxCellFont.Create;
  FFill := TXlsxCellFill.Create;

  FBottomBorder := TXlsxBorderStyle.xbsNone;
  FTopBorder := TXlsxBorderStyle.xbsNone;
  FDiagonalBorder := TXlsxBorderStyle.xbsNone;
  FLeftBorder := TXlsxBorderStyle.xbsNone;
  FRightBorder := TXlsxBorderStyle.xbsNone;
end;

destructor TXlsxCellFormat.Destroy;
begin
  FFont.Free;
  FFill.Free;
  inherited;
end;

function TXlsxCellFormat.IsSame(format: TXlsxCellFormat): Boolean;
begin
  Result := Font.IsSame(format.Font) and
            Fill.IsSame(format.Fill) and
            (HasSameBorders(format));
end;

function TXlsxCellFormat.HasSameBorders(format: TXlsxCellFormat): Boolean;
begin
  Result := (LeftBorder = format.LeftBorder) and
            (RightBorder = format.RightBorder) and
            (TopBorder = format.TopBorder) and
            (BottomBorder = format.BottomBorder) and
            (DiagonalBorder = format.DiagonalBorder);
end;

procedure TXlsxCellFormat.SaveBorder(nodename: String; style: TXlsxBorderStyle; node: TJvSimpleXMLElem);
var
  borderNode: TJvSimpleXMLElem;
begin
  borderNode := node.Items.Add(nodeName);

  if style <> xbsNone then
  begin
    borderNode.Properties.Add('style', XlsxBorderStyleName[style]);
    { TODO : color for border }
//    borderNode.Items.Add('color').Properties.Add('rgb', 'FFFF0000');
  end;
end;

procedure TXlsxCellFormat.SaveToBorderNode(node: TJvSimpleXMLElem);
var
  borderNode: TJvSimpleXMLElem;
begin
  borderNode := node.Items.Add('border');
  SaveBorder('left', LeftBorder, borderNode);
  SaveBorder('right', RightBorder, borderNode);
  SaveBorder('top', TopBorder, borderNode);
  SaveBorder('bottom', BottomBorder, borderNode);
  SaveBorder('diagonal', DiagonalBorder, borderNode);

  { TODO : Style for diagonal border }
  if DiagonalBorder <> xbsNone then
  begin
    borderNode.Properties.Add('diagonalUp', 1);
    borderNode.Properties.Add('diagonalDown', 1);
  end;
end;

procedure TXlsxCellFormat.SaveToFillsNode(node: TJvSimpleXMLElem);
var
  fillNode: TJvSimpleXMLElem;
  patternFillNode: TJvSimpleXMLElem;
  fgColorNode: TJvSimpleXMLElem;
  bgColorNode: TJvSimpleXMLElem;
begin
  fillNode := node.Items.Add('fill');
  patternFillNode := fillNode.Items.Add('patternFill');
  patternFillNode.Properties.Add('patternType', XlsxFillPatternTypeName[Fill.PatternType]);

  if fill.PatternType <> xptNone then
  begin
    fgColorNode := patternFillNode.Items.Add('fgColor');
    fgColorNode.Properties.Add('rgb', Fill.fgToRGB);
    bgColorNode := patternFillNode.Items.Add('bgColor');
    bgColorNode.Properties.Add('rgb', Fill.bgToRGB);
  end;
end;

procedure TXlsxCellFormat.SaveToFontNode(node: TJvSimpleXMLElem);
var
  fontNode: TJvSimpleXMLElem;
begin
  fontNode := node.Items.Add('font');
  if xfsBold in Font.Style then fontNode.Items.Add('b');
  if xfsItalic in Font.Style then fontNode.Items.Add('i');
  if xfsUnderline in Font.Style then fontNode.Items.Add('u');

  fontNode.Items.Add('sz').Properties.Add('val', Font.Size);
  fontNode.Items.Add('name').Properties.Add('val', Font.FontName);
end;

{ TXlsxDistinctFormatList }

function TXlsxDistinctFormatList.AddIfNotExists(
  format: TXlsxCellFormat): Integer;
var
  value: TXlsxCellFormat;
  i: Integer;
  doAdd: Boolean;
begin
  Result := -1;
  doAdd := true;
  for i := 0 to Self.Count-1 do
  begin
    value := Self[i];
    if value.IsSame(format) then
    begin
      doAdd := false;
      Result := i;
      break;
    end;
  end;
  if doAdd then
  begin
    Result := Self.Add(format);
  end;
end;

function TXlsxDistinctFormatList.getFormatById(id: Integer): TXlsxCellFormat;
var
  cellFormat: TXlsxCellFormat;
begin
  Result := nil;
  for cellFormat in Self do
  begin
    if cellFormat.FormatId = id then
    begin
      Result := cellFormat;
      break;
    end;
  end;
end;

{ TXlsxCellFont }

procedure TXlsxCellFont.Assign(value: TXlsxCellFont);
begin
  FColorTheme := value.ColorTheme;
  FSize := value.Size;
  FFontName := value.FontName;
  FFamily := value.Family;
  FStyle := value.Style;
end;

constructor TXlsxCellFont.Create;
begin
  FColorTheme := 1;
  FSize := 11;
  FFontName := 'Calibri';
  FFamily := 2;
  FStyle := [];
  fIndexedColor := -1;
end;

function TXlsxCellFont.IsSame(font: TXlsxCellFont): Boolean;
begin
  Result := (Size = font.Size) and
            (ColorTheme = font.ColorTheme) and
            (FontName = font.FontName) and
            (Family = font.Family) and
            (Style = font.Style);
end;

procedure TXlsxCellFont.SetIndexedColor(idx: Integer);
begin
  fIndexedColor := idx;
end;

{ TXlsxBorder }

constructor TXlsxBorder.Create;
begin
  FStyle := xbsNone;
end;

function TXlsxBorder.getBorderStyleToString: String;
begin
  case FStyle of
    xbsNone: Result := 'none';
    xbsMediumDashDotDot: Result := 'mediumDashDotDot';
    xbsHair: Result := 'hair';
    xbsSlantDashDot: Result := 'slantDashDot';
    xbsDotted: Result := 'dotted';
    xbsMediumDashDot: Result := 'mediumDashDot';
    xbsDashDotDot: Result := 'dashDotDot';
    xbsMediumDashed: Result := 'mediumDashed';
    xbsDashDot: Result := 'dashDot';
    xbsMedium: Result := 'medium';
    xbsDashed: Result := 'dashed';
    xbsThick: Result := 'thick';
    xbsThin: Result := 'thin';
    xbsDouble: Result := 'double';
  end;
end;

procedure TXlsxBorder.SaveToNode(nodeName: String; node: TJvSimpleXMLElem);
var
  borderNode: TJvSimpleXMLElem;
begin
  borderNode := node.Items.Add(nodeName);

  if FStyle <> xbsNone then
  begin
    borderNode.Properties.Add('style', getBorderStyleToString)
  end;
end;

procedure TXlsxBorder.SetPatternTypeByName(StyleName: String);
begin
  if StyleName.Equals('none') then fStyle := xbsNone;
  if StyleName.Equals('mediumDashDotDot') then fStyle := xbsMediumDashDotDot;
  if StyleName.Equals('hair') then fStyle := xbsHair;
  if StyleName.Equals('slantDashDot') then fStyle := xbsSlantDashDot;
  if StyleName.Equals('dotted') then fStyle := xbsDotted;
  if StyleName.Equals('mediumDashDot') then fStyle := xbsMediumDashDot;
  if StyleName.Equals('dashDotDot') then fStyle := xbsDashDotDot;
  if StyleName.Equals('mediumDashed') then fStyle := xbsMediumDashed;
  if StyleName.Equals('dashDot') then fStyle := xbsDashDot;
  if StyleName.Equals('medium') then fStyle := xbsMedium;
  if StyleName.Equals('dashed') then fStyle := xbsDashed;
  if StyleName.Equals('thick') then fStyle := xbsThick;
  if StyleName.Equals('thin') then fStyle := xbsThin;
  if StyleName.Equals('double') then fStyle := xbsDouble;
end;

{ TXlsxCellFill }

procedure TXlsxCellFill.Assign(value: TXlsxCellFill);
begin
  FPatternType := value.PatternType;
  fFgColor := value.fgColor;
  fBgColor := value.bgColor;
end;

function TXlsxCellFill.bgToRGB: String;
var
  tmpColor: TColorRec;
begin
  tmpColor := bgColor;
  Result := tmpColor.GetDelphiHexWithoutDollar;
end;

function TXlsxCellFill.SetbgColorByRGB(rgb: String): String;
var
  tmpColor: TColorRec;
begin
  tmpColor.SetDelphiHexWithoutDollarInput(rgb);
  fBgColor := tmpColor;
end;

function TXlsxCellFill.SetfgColorByRGB(rgb: String): String;
var
  tmpColor: TColorRec;
begin
  tmpColor.SetDelphiHexWithoutDollarInput(rgb);
  fFgColor := tmpColor;
end;

constructor TXlsxCellFill.Create;
begin
  FPatternType := xptNone;
  fFgColor := 0;
  fBgColor := 0;
  fThemeFgColor := 0;
  fThemeBgColor := 0;
  fTintFgColor  := 1;
  fTintBgColor  := 1;
end;

function TXlsxCellFill.IsSame(value: TXlsxCellFill): Boolean;
begin
  Result := (value.PatternType = self.PatternType) and
            (value.fgColor = self.fgColor) and
            (value.bgColor = self.bgColor);
end;

procedure TXlsxCellFill.SetBgIndexedColor(idx: integer);
var
  colorRec: TColorRec;
begin
  colorRec.SetHTMLHexWithoutHashInput(TIndexedColor.Color[idx]);
  bgColor := colorRec;
end;

procedure TXlsxCellFill.SetFgIndexedColor(idx: integer);
var
  colorRec: TColorRec;
begin
  colorRec.SetHTMLHexWithoutHashInput(TIndexedColor.Color[idx]);
  fgColor := colorRec;
end;

procedure TXlsxCellFill.SetPatternTypeByName(PatternName: String);
begin
  if PatternName.Equals('none') then FPatternType := xptNone;
  if PatternName.Equals('gray125') then FPatternType := xptGray125;
  if PatternName.Equals('solid') then FPatternType := xptSolid;
end;

function TXlsxCellFill.fgToRGB: String;
var
  tmpColor: TColorRec;
begin
  tmpColor := fgColor;
  Result := tmpColor.GetDelphiHexWithoutDollar;
end;

end.
