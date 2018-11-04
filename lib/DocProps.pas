unit DocProps;

interface

type
  TXlsxDocumentProperties = class(TObject)
  private
    FCreatedBy: String;
    FApplicationName: String;
    FModifiedBy: String;
    FCreatedAt: TDateTime;
    FModifiedAt: TDateTime;

    procedure InternalSaveApp(filename: String);
    procedure InternalSaveCore(filename: String);
  public
    constructor Create;

    property ApplicationName: String read FApplicationName write FApplicationName;
    property CreatedBy: String read FCreatedBy write FCreatedBy;
    property ModifiedBy: String read FModifiedBy write FModifiedBy;
    property CreatedAt: TDateTime read FCreatedAt write FCreatedAt;
    property ModifiedAt: TDateTime read FModifiedAt write FModifiedAt;

    procedure SaveToXml(basepath: String);
  end;

implementation

uses
  JvSimpleXml, System.IOUtils, System.SysUtils;

{ TXlsxDocumentProperties }

constructor TXlsxDocumentProperties.Create;
begin
  FApplicationName := 'Simple XSLX';
end;

procedure TXlsxDocumentProperties.InternalSaveApp(filename: String);
var
  xmlExport: TJvSimpleXML;
  childNode: TJvSimpleXMLElem;
begin
  xmlExport := TJvSimpleXML.Create(nil);
  childNode := xmlExport.Root.Items.Add('Properties');
  childNode.Properties.Add('xmlns', 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties');
  childNode.Properties.Add('xmlns:vt', 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes');

  childNode.Items.Add('Application', FApplicationName);

  xmlExport.SaveToFile(filename);
  xmlExport.Free;
end;

procedure TXlsxDocumentProperties.InternalSaveCore(filename: String);
var
  xmlExport: TJvSimpleXML;
  childNode: TJvSimpleXMLElem;
begin
  xmlExport := TJvSimpleXML.Create(nil);
  childNode := xmlExport.Root.Items.Add('cp:coreProperties');
  childNode.Properties.Add('xmlns:cp', 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties');
  childNode.Properties.Add('xmlns:dc', 'http://purl.org/dc/elements/1.1/');
  childNode.Properties.Add('xmlns:dcterms', 'http://purl.org/dc/terms/');
  childNode.Properties.Add('xmlns:dcmitype', 'http://purl.org/dc/dcmitype/');
  childNode.Properties.Add('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance');

  childNode.Items.Add('dc:creator', FCreatedBy);
  childNode.Items.Add('cp:lastModifiedBy', FModifiedBy);

//  childNode.Items.Add('dcterms:created', FCreatedAt);
//  childNode.Items.Add('dcterms:modified', FModifiedBy);

  xmlExport.SaveToFile(filename);
  xmlExport.Free;
end;

procedure TXlsxDocumentProperties.SaveToXml(basepath: String);
begin
  basePath := TPath.Combine(basepath, 'docProps');
  ForceDirectories(basePath);
  InternalSaveApp(TPath.Combine(basepath, 'app.xml'));
  InternalSaveCore(TPath.Combine(basepath, 'core.xml'));
end;

end.
