unit Cell;

interface

uses
  JvSimpleXml;

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
    function getReference: String;
  public
    constructor Create(_row, _col: Integer);
    procedure SaveToWorksheetXmlNode(node: TJvSimpleXMLElem);

    property Col: Integer read FCol write FCol;
    property Row: Integer read FRow write FRow;
    property Value: TXlsxCellValue read FValue write FValue;
    property Formula: String read FFormula write FFormula;
  end;

implementation

uses
  Helper, System.SysUtils;

{ TXlsxCell }

constructor TXlsxCell.Create(_row, _col: Integer);
begin
  FCol := _col;
  FRow := _row;
end;

function TXlsxCell.getReference: String;
begin
  Result := ColNumberToReference(Col)+IntToStr(Row);
end;

procedure TXlsxCell.SaveToWorksheetXmlNode(node: TJvSimpleXMLElem);
var
  rowNode: TJvSimpleXMLElem;
  cellNode: TJvSimpleXMLElem;
begin
  rowNode := node.Items.Add('row');
  rowNode.Properties.Add('r', Row);

  cellNode := rowNode.Items.Add('c');
  cellNode.Properties.Add('r', getReference);

  if FFormula <> '' then
  begin
    cellNode.Items.Add('f', FFormula);
  end else
  begin
    if value.IsString then
    begin
      cellNode.Properties.Add('t', 'inlineStr');
      cellNode.Items.Add('is').Items.Add('t', Value.AsString);
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
  Result := FStringValue;
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
