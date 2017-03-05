unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, ComObj, StdCtrls, Mask, ExtCtrls, ComCtrls;

type
  TForm1 = class(TForm)
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    MaskEdit1: TMaskEdit;
    Label1: TLabel;
    SG1: TStringGrid;
    OpenDialog1: TOpenDialog;
    Button4: TButton;
    PB1: TProgressBar;
    Label2: TLabel;
    Label3: TLabel;
    Imz1: TImage;
    Imk1: TImage;
    Imz2: TImage;
    Imk2: TImage;
    Imz3: TImage;
    Imk3: TImage;
    Imz4: TImage;
    Imk4: TImage;
    SG2: TStringGrid;
    Button5: TButton;
    Label4: TLabel;
    Label5: TLabel;
    procedure FormShow(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
Const
  xlCellTypeLastCell = $000000B;
var
  d: TDateTime;
  ExLApp, Sheet : OLEVariant;
  i, j, r, c, q:integer;
  color1, color2 , ogid1, ogid2, kod, razdel,p,prise :string;
  kurs, temp, tempr:real;
begin
  Label3.Visible:=true;
  d:=now;
  PB1.Visible:=true;
  SG1.Visible:=true;

  if OpenDialog1.Execute then prise:=OpenDialog1.FileName;
  ExLApp:=CreateOleObject('Excel.Application');
  ExLApp.Visible:=false;
  ExLApp.Workbooks.Open(prise);
  Sheet:=ExLApp.Workbooks[ExtractFileName(prise)].WorkSheets[1];
  Sheet.Cells.SpecialCells(xlCellTypeLastCell).Activate;
    r:=ExLApp.ActiveCell.Row;
    c:=4;
    q:=1;
  PB1.Max:=r;
  SG1.ColCount:=c;

  color1:=sheet.cells[5,5].Interior.Color;  // ���� ��� �����
  color2:=sheet.cells[12,1].Interior.Color; // ���� �������

  kod:=VarToStr(sheet.cells[13,1].NumberFormat);
  kurs:=StrtoFloat(MaskEdit1.Text);
  i:=0;
  for j := 0 to r do      // ������
    Begin

      if sheet.cells[12,2].NumberFormat=sheet.cells[j+12,2].NumberFormat then
        Begin
             ogid1:=sheet.cells[j+12,2];
             if ((ogid1='HD �����������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='USB Flash') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='USB-���� �������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='����� �������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='������� �������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='���-������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='�����������, �������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='����������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='���������������, ������, ��������, GSM ������������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='����������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='�������� ����������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='���, �������������, �������������� �������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='������� ������������, ���������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='������� ���������, ����������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='������� ����������������, ������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='��������� ���������, ����������, ����������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='����������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='����� ������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='���������� � ���������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='�������, ����������� ������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='������������ ���������� (USB-����������)') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='������������, �����������, ����������������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='����������� USB, PCI, PCMCIA') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='�������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='����������� �����') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='���������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='��������� ��������, ���, ���������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='������, �������, ������� �����') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='��������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='���������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='���, ������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='����, �������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='��������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='��������, ����������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='���������� � ���������� ����������, Apple') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='�������� (���������� ����������)') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='�������� �����������, ������������� �����') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='������ DVD, MP3, ���������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='�������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='��������, ��������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='����������� �����������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='���������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='����������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='��������� ���������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='�������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='����������e ������������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='����������, ���������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='��������, �����, ���') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='������� �.�.') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='������������') and (sheet.cells[j+12,1].Interior.Color=color2))
             or ((ogid1='����������� �����, �����������') and (sheet.cells[j+12,1].Interior.Color=color2))
             then
             i:=0
             Else
             Begin
                if (i=0)and(sheet.cells[j+12,1].Interior.Color<>color2) then i:=0
                Else i:=1;

             End;
        End;

      if i=0 then
        Begin
  {  if (sheet.cells[j+12,1].Interior.Color=color1) or (sheet.cells[j+12,1].Interior.Color=color2) then
      Begin  }
              if (sheet.cells[1,1].NumberFormat<>sheet.cells[j+12,3].NumberFormat) and
              (sheet.cells[j+12,1].Interior.Color=color1) then
                Begin
                SG1.Cells[0,q]:=sheet.cells[j+12,1];
                SG1.Cells[1,q]:=sheet.cells[j+12,2];
                SG1.Cells[2,q]:=sheet.cells[j+12,3];
                SG1.Cells[3,q]:=sheet.cells[j+12,4];
                SG1.RowCount:=q+1;
                q:=q+1;
                End;
              if (sheet.cells[1,1].NumberFormat=sheet.cells[j+12,3].NumberFormat)  then
                Begin
                  ogid2:= sheet.cells[j+12,1];
                  if (sheet.cells[12,2].NumberFormat=sheet.cells[j+12,2].NumberFormat) and
                  (ogid2='')then
                    Begin
                      SG1.Cells[1,q]:=sheet.cells[j+12,2];
                      SG1.RowCount:=q+1;
                      q:=q+1;
                    End;
                  if (sheet.cells[1,1].NumberFormat<>sheet.cells[j+12,4].NumberFormat) and
                  (sheet.cells[j+12,1].Interior.Color=color1) then
                    Begin
                      temp:=strtofloat(sheet.cells[j+12,4]);
                      tempr:=temp/kurs;
                      SG1.Cells[0,q]:=sheet.cells[j+12,1];
                      SG1.Cells[1,q]:=sheet.cells[j+12,2];
                      SG1.Cells[2,q]:=FormatFloat('0.##', tempr);
                      SG1.Cells[3,q]:=sheet.cells[j+12,4];
                      SG1.RowCount:=q+1;
                      q:=q+1;
                    End;
                End;
        End;
      PB1.Position:=j;
    End;

  if not VarIsEmpty(ExLApp) then
  begin
    ExLApp.DisplayAlerts := False; // <---
    ExLApp.Quit;
    ExLApp:=Unassigned;
  end;

  PB1.Visible:=false;
  Label5.Visible:=true;
  Label5.Caption:='����� ��������� '+FormatDateTime('hh:mm:ss:zzz', Now()-d)+' ����������� �����:'+IntToStr(SG1.RowCount);
  Label3.Visible:=false;

  Button2.Enabled:=true;
end;

procedure TForm1.Button2Click(Sender: TObject);
Const
  xlCellTypeLastCell = $000000B;
var
  ArrayDate:Variant;
  ExlApp, Range, Sheet: OLEVariant;
  i, j, x, y:integer;
  s,u:string;
  d: TDateTime;
begin
  Label5.Visible:=false;

  d:=now;
  PB1.Max:=SG1.RowCount;
  PB1.Visible:=true;
  PB1.Position:=0;

  ExLApp:=CreateOleObject('Excel.Application');
  ExLApp.Application.WorkBooks.Add();
  Sheet:=ExLApp.Workbooks[1].Worksheets[1];
  for i:=1 to  SG1.RowCount-1 do
      Begin
         sheet.cells[i,1]:=SG1.Cells[0,i];
         sheet.cells[i,2]:=SG1.Cells[1,i];
         s:=SG1.Cells[2,i];
         if s<>'' then sheet.cells[i,3]:=strtofloat(SG1.Cells[2,i]);
         sheet.cells[i,4]:=SG1.Cells[3,i];
         PB1.Position:=i;
      End;
  PB1.Position:=SG1.RowCount;
  PB1.Visible:=false;
  Label5.Visible:=true;
  Label5.Caption:='����� ������ StringGrid '+FormatDateTime('hh:mm:ss:zzz', Now()-d)+' ����������� �����:'+IntToStr(SG1.RowCount);

  ExLApp.Visible:=true;
end;

procedure TForm1.Button5Click(Sender: TObject);

begin
  Application.Terminate;
end;

procedure TForm1.FormShow(Sender: TObject);
begin
  Label3.Visible:=false;
  Label5.Visible:=false;
  PB1.Visible:=false;
  SG1.Visible:=false;
  SG2.Visible:=false;
  // ����� SG1 � SG2
  Sg1.Cells[0,0]:='���';
  Sg1.Cells[1,0]:='������ ��� ������������';
  Sg1.Cells[2,0]:='$';
  Sg1.Cells[3,0]:='���.';

  Sg2.Cells[0,0]:='���';
  Sg2.Cells[1,0]:='������ ��� ������������';
  Sg2.Cells[2,0]:='$';
  Sg2.Cells[3,0]:='���.';

  // ������ ��������� ���������
  Imk1.Visible:=false;
  Imk2.Visible:=false;
  Imk3.Visible:=false;
  Imk4.Visible:=false;
  Imz1.Visible:=false;
  Imz2.Visible:=false;
  Imz3.Visible:=false;
  Imz4.Visible:=false;

  // ������ ����������� ������
  Button2.Enabled:=false;
  Button3.Enabled:=false;
  Button4.Enabled:=false;

end;

end.
