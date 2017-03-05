unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, StdCtrls, Mask, ExtCtrls, ComCtrls;

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
    procedure FormShow(Sender: TObject);
    procedure Button5Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.Button5Click(Sender: TObject);
begin
  Application.Terminate;
end;

procedure TForm1.FormShow(Sender: TObject);
begin
  Label3.Visible:=false;
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
