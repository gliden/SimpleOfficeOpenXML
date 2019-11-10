unit IndexedColors;

interface

uses
  System.Generics.Collections;

type
  TIndexedColor = class(TObject)
  private
    class var fColors: TDictionary<Integer, String>;
    class function getColor(index: Integer): String; static;
  public
    class constructor Create;
    class destructor Destroy;

    class property Color[index: Integer]: String read getColor;
  end;

implementation

{ TIndexedColor }

class constructor TIndexedColor.Create;
begin
  fColors := TDictionary<Integer, String>.Create;

  fColors.Add(0, '000000');
  fColors.Add(1, 'FFFFFF');
  fColors.Add(2, 'FF0000');
  fColors.Add(3, '00FF00');
  fColors.Add(4, '0000FF');
  fColors.Add(5, 'FFFF00');
  fColors.Add(6, 'FF00FF');
  fColors.Add(7, '00FFFF');
  fColors.Add(8, '000000');
  fColors.Add(9, 'FFFFFF');
  fColors.Add(10, 'FF0000');
  fColors.Add(11, '00FF00');
  fColors.Add(12, '0000FF');
  fColors.Add(13, 'FFFF00');
  fColors.Add(14, 'FF00FF');
  fColors.Add(15, '00FFFF');
  fColors.Add(16, '800000');
  fColors.Add(17, '008000');
  fColors.Add(18, '000080');
  fColors.Add(19, '808000');
  fColors.Add(20, '800080');
  fColors.Add(21, '008080');
  fColors.Add(22, 'C0C0C0');
  fColors.Add(23, '808080');
  fColors.Add(24, '9999FF');
  fColors.Add(25, '993366');
  fColors.Add(26, 'FFFFCC');
  fColors.Add(27, 'CCFFFF');
  fColors.Add(28, '660066');
  fColors.Add(29, 'FF8080');
  fColors.Add(30, '0066CC');
  fColors.Add(31, 'CCCCFF');
  fColors.Add(32, '000080');
  fColors.Add(33, 'FF00FF');
  fColors.Add(34, 'FFFF00');
  fColors.Add(35, '00FFFF');
  fColors.Add(36, '800080');
  fColors.Add(37, '800000');
  fColors.Add(38, '008080');
  fColors.Add(39, '0000FF');
  fColors.Add(40, '00CCFF');
  fColors.Add(41, 'CCFFFF');
  fColors.Add(42, 'CCFFCC');
  fColors.Add(43, 'FFFF99');
  fColors.Add(44, '99CCFF');
  fColors.Add(45, 'FF99CC');
  fColors.Add(46, 'CC99FF');
  fColors.Add(47, 'FFCC99');
  fColors.Add(48, '3366FF');
  fColors.Add(49, '33CCCC');
  fColors.Add(50, '99CC00');
  fColors.Add(51, 'FFCC00');
  fColors.Add(52, 'FF9900');
  fColors.Add(53, 'FF6600');
  fColors.Add(54, '666699');
  fColors.Add(55, '969696');
  fColors.Add(56, '003366');
  fColors.Add(57, '339966');
  fColors.Add(58, '003300');
  fColors.Add(59, '333300');
  fColors.Add(60, '993300');
  fColors.Add(61, '993366');
  fColors.Add(62, '333399');
  fColors.Add(63, '333333');
end;

class destructor TIndexedColor.Destroy;
begin
  fColors.Free;
  inherited;
end;

class function TIndexedColor.getColor(index: Integer): String;
begin
  if not fColors.TryGetValue(index, Result) then Result := '';
end;

end.
