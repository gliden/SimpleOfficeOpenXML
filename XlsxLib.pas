unit XlsxLib;

interface

uses
  DocProps, Workbook, System.Types, StylesFile;

type
  TXlsxFile = class(TObject)
  private
    FDocumentProperties: TXlsxDocumentProperties;
    FWorkbook: TXlsxWorkbook;
    procedure SaveToXml(basePath: String);
  public
    constructor Create;
    destructor Destroy;override;

    property DocumentProperties: TXlsxDocumentProperties read FDocumentProperties;
    property Workbook: TXlsxWorkbook read FWorkbook;
    procedure SaveToFile(filename: String);
    procedure LoadFromFile(filename: String);
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

procedure TXlsxFile.LoadFromFile(filename: String);
var
  basePath: string;
begin
  FWorkbook.Sheets.Clear;
  basePath := TPath.Combine(TPath.GetTempPath, TGuid.NewGuid.ToString);
  ForceDirectories(basePath);
  TZipFile.ExtractZipFile(filename, basePath);

  FWorkbook.LoadFromXml(basePath);


  TDirectory.Delete(basePath, true);
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
