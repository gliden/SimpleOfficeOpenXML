unit XlsxLib;

interface

uses
  DocProps, Workbook, System.Types, StylesFile;

type
  TXlsxFile = class(TObject)
  private
    FDocumentProperties: TXlsxDocumentProperties;
    FWorkbook: TXlsxWorkbook;
  public
    constructor Create;
    destructor Destroy;override;

    property DocumentProperties: TXlsxDocumentProperties read FDocumentProperties;
    property Workbook: TXlsxWorkbook read FWorkbook;
    procedure SaveToXml(basePath: String);
    procedure SaveToFile(filename: String);
  end;

implementation

uses
  ContentTypes, System.SysUtils, Rels, System.IOUtils, System.Zip;

{ TXlsxFile }

constructor TXlsxFile.Create;
begin
  FDocumentProperties := TXlsxDocumentProperties.Create;
  FWorkbook := TXlsxWorkbook.Create;
  FWorkbook.AddSheet('Sheet 1');
end;

destructor TXlsxFile.Destroy;
begin
  FDocumentProperties.Free;
  FWorkbook.Free;
  inherited;
end;

procedure TXlsxFile.SaveToFile(filename: String);
var
  files: TStringDynArray;
  zipFile: TZipFile;
  tmpFilename: String;
  s: string;
  basePath: string;
begin
  if TFile.Exists(filename) then TFile.Delete(filename);
  basePath := TPath.Combine(TPath.GetTempPath, TGuid.NewGuid.ToString);
  ForceDirectories(basePath);
  SaveToXml(basePath);

  files := TDirectory.GetFiles(basePath, '*.*', TSearchOption.soAllDirectories);

  zipFile := TZipFile.Create;
  zipFile.Open(filename, TZipMode.zmWrite);
  for tmpFilename in files do
  begin
    s := tmpFileName.Replace(basePath, '');
    if s.StartsWith(PathDelim) then s := s.Substring(1);

    zipFile.Add(tmpFilename, s);
  end;
  zipFile.Free;

  TDirectory.Delete(basePath, true);
end;

procedure TXlsxFile.SaveToXml(basePath: String);
var
  contentType: TXlsxContentTypes;
  globalRels: TXlsxRels;
  fStyles: TXlsxStylesFile;
begin
  fStyles := TXlsxStylesFile.Create;
  fStyles.BuildFormatList(FWorkbook);
  fStyles.SaveToXml(basePath);
  fStyles.Free;

  contentType := TXlsxContentTypes.Create;
  contentType.SaveToXml(basePath);
  contentType.Free;

  globalRels := TXlsxRels.Create;
  globalRels.SaveToXml(basePath);
  globalRels.Free;

  FDocumentProperties.SaveToXml(basePath);
  FWorkbook.SaveToXml(basePath);
end;

end.
