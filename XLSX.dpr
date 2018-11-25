program XLSX;

uses
  Vcl.Forms,
  XlsxFrm in 'XlsxFrm.pas' {XlsxDlg},
  XlsxLib in 'XlsxLib.pas',
  ContentTypes in 'lib\ContentTypes.pas',
  Rels in 'lib\Rels.pas',
  DocProps in 'lib\DocProps.pas',
  Workbook in 'lib\Workbook.pas',
  Sheets in 'lib\Sheets.pas',
  Cell in 'lib\Cell.pas',
  Helper in 'lib\Helper.pas',
  CellFormat in 'lib\CellFormat.pas',
  StylesFile in 'lib\StylesFile.pas',
  SharedStrings in 'lib\SharedStrings.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TXlsxDlg, XlsxDlg);
  Application.Run;
end.
