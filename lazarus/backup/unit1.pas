unit Unit1;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, ActnList;

type

  { TForm1 }

  TForm1 = class(TForm)
    StartCalc: TButton;
    SelectExcelFileAct: TAction;
    SaveExcelFileAct: TAction;
    ActionList1: TActionList;
    SelectExcelFileBtn: TButton;
    SaveExcelFileBtn: TButton;
    SelectExcelFileEdit: TEdit;
    SaveExcelFileEdit: TEdit;
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    GroupBox3: TGroupBox;
    OutputMemo: TMemo;
    SelectExcelFileDia: TOpenDialog;
    SaveExcelFileDia: TSaveDialog;
    procedure SaveExcelFileActExecute(Sender: TObject);
    procedure SelectExcelFileActExecute(Sender: TObject);
    procedure SelectExcelFileBtnClick(Sender: TObject);
    procedure SaveExcelFileBtnClick(Sender: TObject);
    procedure StartCalcClick(Sender: TObject);
  private

  public

  end;

var
  Form1: TForm1;

implementation

{$R *.lfm}

{ TForm1 }

procedure TForm1.SelectExcelFileBtnClick(Sender: TObject);
begin

end;

procedure TForm1.SelectExcelFileActExecute(Sender: TObject);
begin

end;

procedure TForm1.SaveExcelFileActExecute(Sender: TObject);
begin

end;

procedure TForm1.SaveExcelFileBtnClick(Sender: TObject);
begin

end;

procedure TForm1.StartCalcClick(Sender: TObject);
begin

end;

end.

