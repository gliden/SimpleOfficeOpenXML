unit Cell;

interface

uses
  JvSimpleXml, CellFormat, SharedStrings;

type
  TXlsxCellValueType = (vtString, vtNumber, vtBoolean);
  TXlsxCellValue = record
  private
    FCellValueType: TXlsxCellValueType;
    FStringValue: String;
    FNumberValue: Double;
    FBoolValue: Boolean;
    function AsXlsxBoolean: Integer;
  public
    class operator Implicit(const value: String): TXlsxCellValue;
    class operator Implicit(const value: Double): TXlsxCellValue;
    class operator Implicit(const value: Boolean): TXlsxCellValue;

    function IsString: Boolean;
    function AsString: String;

    function IsNumber: Boolean;
    function AsFloat: Double;

    function IsBoolean: Boolean;
    function AsBoolean: Boolean;
  end;

  TXlsxCell = class(TObject)
  private
    FCol: Integer;
    FRow: Integer;
    FValue: TXlsxCellValue;
    FFormula: String;
    FFormat: TXlsxCellFormat;
    FSharedStrings: TXlsxSharedStrings;
    function getReference: String;
    function GetRowNode(node: TJvSimpleXMLElem; row: Integer): TJvSimpleXMLElem;
  public
    constructor Create(_row, _col: Integer; cellFormat: TXlsxCellFormat; sharedStrings: TXlsxSharedStrings);
    destructor Destroy;override;
    procedure SaveToWorksheetXmlNode(node: TJvSimpleXMLElem);

    property Col: Integer read FCol write FCol;
    property Row: Integer read FRow write FRow;
    property Value: TXlsxCellValue read FValue write FValue;
    property Formula: String read FFormula write FFormula;

    property Format: TXlsxCellFormat read FFormat;
  end;

implementation

uses
  Helper, System.SysUtils;

{ TXlsxCell }

constructor TXlsxCell.Create(_row, _col: Integer; cellFormat: TXlsxCellFormat; sharedStrings: TXlsxSharedStrings);
begin
  FFormat := TXlsxCellFormat.Create;
  if cellFormat <> nil then FFormat.Assign(cellFormat);
  FSharedStrings := sharedStrings;

  FCol := _col;
  FRow := _row;
end;

destructor TXlsxCell.Destroy;
begin
  Format.Free;
  inherited;
end;

function TXlsxCell.GetRowNode(node: TJvSimpleXMLElem; row: Integer): TJvSimpleXMLElem;
var
  i: Integer;
  tmpNode: TJvSimpleXMLElem;
begin
  Result := nil;
  for i := 0 to node.ItemCount-1 do
  begin
    tmpNode := node.Items[i];
    if SameText(tmpNode.Name, 'row') and (tmpNode.Properties.IntValue('r') = row) then
    begin
      Result := tmpNode;
      break;
    end;
  end;
  if Result = nil then
  begin
    Result := node.Items.Add('row');
    Result.Properties.Add('r', row);
  end;
end;

function TXlsxCell.getReference: String;
begin
  Result := ColNumberToReference(Col)+IntToStr(Row);
end;

procedure TXlsxCell.SaveToWorksheetXmlNode(node: TJvSimpleXMLElem);
var
  rowNode: TJvSimpleXMLElem;
  cellNode: TJvSimpleXMLElem;
  sharedStringId: Integer;
begin
  rowNode := GetRowNode(node, row);

  cellNode := rowNode.Items.Add('c');
  cellNode.Properties.Add('r', getReference);

  cellNode.Properties.Add('s', FFormat.FormatId);

  if FFormula <> '' then
  begin
    cellNode.Items.Add('f', FFormula);
  end else
  begin
    if value.IsString then
    begin
      sharedStringId := FSharedStrings.AddString(value.AsString);
      cellNode.Properties.Add('t', 's');
      cellNode.Items.Add('v', sharedStringId);
    end else
    if value.IsNumber then
    begin
      cellNode.Properties.Add('t', 'n');
      cellNode.Items.Add('v').FloatValue := Value.AsFloat;
    end else
    if value.IsBoolean then
    begin
      cellNode.Properties.Add('t', 'b');
      cellNode.Items.Add('v').IntValue := Value.AsXlsxBoolean;
    end;
  end;
end;

{ TXlsxCellValue }

class operator TXlsxCellValue.Implicit(const value: String): TXlsxCellValue;
begin
  Result.FCellValueType := TXlsxCellValueType.vtString;
  Result.FStringValue := value;
end;

function TXlsxCellValue.AsBoolean: Boolean;
begin
  Result := FBoolValue;
end;

function TXlsxCellValue.AsFloat: Double;
begin
  Result := FNumberValue;
end;

function TXlsxCellValue.AsString: String;
begin
  if IsNumber then
  begin
    Result := FloatToStr(FNumberValue);
  end else
  if IsBoolean then
  begin
    Result := BoolToStr(FBoolValue, true);
  end else
  begin
    Result := FStringValue;
  end;
end;

function TXlsxCellValue.AsXlsxBoolean: Integer;
begin
  if FBoolValue then Result := 1 else Result := 0;
end;

class operator TXlsxCellValue.Implicit(const value: Double): TXlsxCellValue;
begin
  Result.FCellValueType := TXlsxCellValueType.vtNumber;
  Result.FNumberValue := value;
end;

class operator TXlsxCellValue.Implicit(const value: Boolean): TXlsxCellValue;
begin
  Result.FCellValueType := TXlsxCellValueType.vtBoolean;
  Result.FBoolValue := value;
end;

function TXlsxCellValue.IsBoolean: Boolean;
begin
  Result := FCellValueType = TXlsxCellValueType.vtBoolean;
end;

function TXlsxCellValue.IsNumber: Boolean;
begin
  Result := FCellValueType = TXlsxCellValueType.vtNumber;
end;

function TXlsxCellValue.IsString: Boolean;
begin
  Result := FCellValueType = TXlsxCellValueType.vtString;
end;

end.
