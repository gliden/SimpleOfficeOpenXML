unit Rels;

interface

type
  TXlsxRels = class(TObject)
  private
  public
    procedure SaveToXml(basepath: String);
  end;

implementation

uses
  JvSimpleXml, System.SysUtils, System.IOUtils;

{ TXlsxRels }

procedure TXlsxRels.SaveToXml(basepath: String);
var
  xmlExport: TJvSimpleXML;
  filename: String;
  relationshipsNode: TJvSimpleXMLElem;
  relationshipNode: TJvSimpleXMLElem;
begin
  filename := TPath.Combine(basepath, '_rels');
  ForceDirectories(filename);
  filename := TPath.Combine(filename, '.rels');

  xmlExport := TJvSimpleXML.Create(nil);
  relationshipsNode := xmlExport.Root.Items.Add('Relationships');
  relationshipsNode.Properties.Add('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

  relationshipNode := relationshipsNode.Items.Add('Relationship');
  relationshipNode.Properties.Add('Id', 'rId3');
  relationshipNode.Properties.Add('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties');
  relationshipNode.Properties.Add('Target', 'docProps/app.xml');

  relationshipNode := relationshipsNode.Items.Add('Relationship');
  relationshipNode.Properties.Add('Id', 'rId2');
  relationshipNode.Properties.Add('Type', 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties');
  relationshipNode.Properties.Add('Target', 'docProps/core.xml');

  relationshipNode := relationshipsNode.Items.Add('Relationship');
  relationshipNode.Properties.Add('Id', 'rId1');
  relationshipNode.Properties.Add('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument');
  relationshipNode.Properties.Add('Target', 'xl/workbook.xml');

  xmlExport.SaveToFile(filename);
  xmlExport.Free;
end;

end.
