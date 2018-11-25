unit ContentTypes;

interface

uses
  JvSimpleXml;

type
  TXlsxContentTypes = class(TObject)
  private
  public
    procedure SaveToXml(basepath: String);
  end;

implementation

uses
  System.IOUtils, JclStreams;

{ TXlsxContentTypes }

procedure TXlsxContentTypes.SaveToXml(basepath: String);
var
  xmlExport: TJvSimpleXML;
  filename: string;
  typeNode: TJvSimpleXMLElem;
  childNode: TJvSimpleXMLElem;
begin
  filename := TPath.Combine(basepath, '[Content_Types].xml');
  xmlExport := TJvSimpleXML.Create(nil);
  typeNode := xmlExport.Root.Items.Add('Types');
  typeNode.Properties.Add('xmlns', 'http://schemas.openxmlformats.org/package/2006/content-types');

  childNode := typeNode.Items.Add('Override');
  childNode.Properties.Add('PartName', '/xl/styles.xml');
  childNode.Properties.Add('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml');

  childNode := typeNode.Items.Add('Default');
  childNode.Properties.Add('Extension', 'rels');
  childNode.Properties.Add('ContentType', 'application/vnd.openxmlformats-package.relationships+xml');

  childNode := typeNode.Items.Add('Default');
  childNode.Properties.Add('Extension', 'xml');
  childNode.Properties.Add('ContentType', 'application/xml');

  childNode := typeNode.Items.Add('Override');
  childNode.Properties.Add('PartName', '/xl/workbook.xml');
  childNode.Properties.Add('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml');

  childNode := typeNode.Items.Add('Override');
  childNode.Properties.Add('PartName', '/docProps/app.xml');
  childNode.Properties.Add('ContentType', 'application/vnd.openxmlformats-officedocument.extended-properties+xml');

  childNode := typeNode.Items.Add('Override');
  childNode.Properties.Add('PartName', '/xl/worksheets/sheet1.xml');
  childNode.Properties.Add('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml');

  childNode := typeNode.Items.Add('Override');
  childNode.Properties.Add('PartName', '/docProps/core.xml');
  childNode.Properties.Add('ContentType', 'application/vnd.openxmlformats-package.core-properties+xml');

  childNode := typeNode.Items.Add('Override');
  childNode.Properties.Add('PartName', '/xl/sharedStrings.xml');
  childNode.Properties.Add('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml');

  xmlExport.SaveToFile(filename, TJclStringEncoding.seUTF8);
  xmlExport.Free;
end;

end.
