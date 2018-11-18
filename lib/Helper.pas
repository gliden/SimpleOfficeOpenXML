unit Helper;

interface

uses
  JvSimpleXML, System.Classes, System.SysUtils;

function ColNumberToReference(col: Integer): String;
procedure SetNameSpace(node: TJvSimpleXMLElem; namespace: String);

implementation

uses
  JclStreams;

function ColNumberToReference(col: Integer): String;
begin
  Assert(col<=26, 'More than 26 cols are currently not supported');
  Result := Char(65+col -1);
end;

procedure SetNameSpace(node: TJvSimpleXMLElem; namespace: String);
var
  i: Integer;
  subNode: TJvSimpleXMLElem;
begin
  node.NameSpace := namespace;
  for i := 0 to node.ItemCount-1 do
  begin
    subNode := node.Items[i];
    SetNameSpace(subNode, namespace);
  end;
end;

end.
