unit Helper;

interface

function ColNumberToReference(col: Integer): String;

implementation

function ColNumberToReference(col: Integer): String;
begin
  Assert(col<=26, 'More than 26 cols are currently not supported');
  Result := Char(65+col -1);
end;

end.
