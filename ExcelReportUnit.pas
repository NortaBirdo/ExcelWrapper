unit ExcelReportUnit;

interface

uses ComObj, System.Classes, Vcl.Dialogs;

type

  TExcelReport = class(TObject)
   private
   var
    xlsBooks: variant;
    ActiveBook: integer;
    ActiveSheet: integer;
   public
    procedure SetActiveBook(id: integer);
    procedure SetActiveSheet(id: integer);
    procedure OpenBooks(FileName: string; Visible: boolean);
    procedure CreateBook(FileName: string; Visible: boolean);
    procedure CloseBooks;
    function ReadCell(Row, Col: integer): string;
    function ReadRow(Row, untilCol:integer ): TStringList;
    procedure WriteCell(Row, Col: integer; Data: string);
    procedure WriteRow(Row, Col: integer; Data: TStringList);
    procedure SetColumnWidth(width: integer; Col: string);
  end;

implementation

{ TExcelReport }

procedure TExcelReport.CloseBooks;
begin
 xlsBooks.Workbooks.Close;
 xlsBooks.Application.quit;
 xlsBooks := 0;
end;

procedure TExcelReport.OpenBooks;
begin
 try
   xlsBooks := CreateOleObject('Excel.Application');
   xlsBooks.WorkBooks.Open(FileName);
   xlsBooks.Visible := Visible;
 except
   ShowMessage('Не надена рабочая копия Excel.');
 end;
end;

procedure TExcelReport.CreateBook(FileName: string; Visible: boolean);
begin
 try
   xlsBooks := CreateOleObject('Excel.Application');
   xlsBooks.Workbooks.Add;
   xlsBooks.Visible := Visible;
 except
   ShowMessage('Не надена рабочая копия Excel.');
 end;
end;

function TExcelReport.ReadCell(Row, Col: integer): string;
begin
 result := xlsBooks.WorkBooks[ActiveBook].WorkSheets[ActiveSheet].Cells[Row, Col];
end;

function TExcelReport.ReadRow(Row,
  untilCol: integer): TStringList;
var
  i: integer;
begin
  result := TStringList.Create;
  for I := 1 to untilCol do
    result.Add(ReadCell(Row, I));
end;

procedure TExcelReport.SetActiveBook(id: integer);
begin
  ActiveBook := id;
end;

procedure TExcelReport.SetActiveSheet(id: integer);
begin
  ActiveSheet := id;
end;

procedure TExcelReport.SetColumnWidth(width: integer; Col: string);
var
  Sheet: OLEVariant;
begin
  Sheet:=xlsBooks.ActiveWorkbook.ActiveSheet;
  sheet.Columns.Range[Col].ColumnWidth := width;
end;

procedure TExcelReport.WriteCell(Row, Col: integer; Data: string);
begin
  xlsBooks.WorkBooks[ActiveBook].WorkSheets[ActiveSheet].Cells[Row, Col] := Data;
end;

procedure TExcelReport.WriteRow(Row, Col: integer; Data: TStringList);
var
  i: integer;
begin
  for I := 0 to Data.Count-1 do
    WriteCell(Row, i, Data.Strings[i]);
end;

end.
