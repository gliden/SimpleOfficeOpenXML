unit XlsxFrm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, XlsxLib;

type
  TXlsxDlg = class(TForm)
    Button1: TButton;
    Button2: TButton;
    procedure Button1Click(Sender: TObject);
  private
    { Private-Deklarationen }
  public
    { Public-Deklarationen }
  end;

var
  XlsxDlg: TXlsxDlg;

implementation

{$R *.dfm}

procedure TXlsxDlg.Button1Click(Sender: TObject);
var
  xlsxFile: TXlsxFile;
begin
  xlsxFile := TXlsxFile.Create;
  xlsxFile.Workbook.Sheets[0].Cell[1, 1].Value := 1;
//  xlsxFile.Workbook.Sheets[0].Cell[2, 1].Value := 2;
//  xlsxFile.Workbook.Sheets[0].Cell[3, 1].Formula := 'A1+B1';

  xlsxFile.SaveToFile('C:\temp\mappe\test.xlsx');
  xlsxFile.Free;
end;

end.
