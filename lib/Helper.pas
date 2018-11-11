unit Helper;

interface

uses
  JvSimpleXML, System.Classes, System.SysUtils;

function ColNumberToReference(col: Integer): String;

implementation

uses
  JclStreams;

function ColNumberToReference(col: Integer): String;
begin
  Assert(col<=26, 'More than 26 cols are currently not supported');
  Result := Char(65+col -1);
end;

end.
