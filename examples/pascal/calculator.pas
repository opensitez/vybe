{ Pascal Calculator with VCL/GCL-style GUI — mirrors examples/vb/calculator.vb
  Run: vybex examples/pascal/calculator.pas }
program Calculator;

uses Forms, StdCtrls;

type
  TCalcForm = class(TForm)
    display: TEdit;
    btn7, btn8, btn9, btnDiv: TButton;
    btn4, btn5, btn6, btnMul: TButton;
    btn1, btn2, btn3, btnSub: TButton;
    btnC, btn0, btnEq, btnAdd: TButton;
    FCurrent: string;
    FPrevious: string;
    FOp: string;
    FResetNext: Boolean;
    constructor Create(AOwner: TObject); override;
    procedure UpdateDisplay;
    procedure PressDigit(d: string);
    procedure PressOperator(op: string);
    procedure DoCalculate;
    procedure PressClear;
    procedure Btn7Click(Sender: TObject);
    procedure Btn8Click(Sender: TObject);
    procedure Btn9Click(Sender: TObject);
    procedure BtnDivClick(Sender: TObject);
    procedure Btn4Click(Sender: TObject);
    procedure Btn5Click(Sender: TObject);
    procedure Btn6Click(Sender: TObject);
    procedure BtnMulClick(Sender: TObject);
    procedure Btn1Click(Sender: TObject);
    procedure Btn2Click(Sender: TObject);
    procedure Btn3Click(Sender: TObject);
    procedure BtnSubClick(Sender: TObject);
    procedure BtnCClick(Sender: TObject);
    procedure Btn0Click(Sender: TObject);
    procedure BtnEqClick(Sender: TObject);
    procedure BtnAddClick(Sender: TObject);
  end;

constructor TCalcForm.Create(AOwner: TObject);
begin
  inherited Create(AOwner);
  Caption := 'Calculator';
  Width := 280;
  Height := 400;

  FCurrent := '0';
  FPrevious := '';
  FOp := '';
  FResetNext := False;

  { Display }
  display := TEdit.Create(Self);
  display.Name := 'display';
  display.Text := '0';
  display.Left := 10;
  display.Top := 10;
  display.Width := 250;
  display.Height := 40;
  display.ReadOnly := True;
  Self.Controls.Add(display);

  { Row 1: 7 8 9 / }
  btn7 := TButton.Create(Self);
  btn7.Name := 'btn7';
  btn7.Caption := '7';
  btn7.Left := 10;
  btn7.Top := 60;
  btn7.Width := 58;
  btn7.Height := 48;
  btn7.OnClick := Btn7Click;
  Self.Controls.Add(btn7);

  btn8 := TButton.Create(Self);
  btn8.Name := 'btn8';
  btn8.Caption := '8';
  btn8.Left := 73;
  btn8.Top := 60;
  btn8.Width := 58;
  btn8.Height := 48;
  btn8.OnClick := Btn8Click;
  Self.Controls.Add(btn8);

  btn9 := TButton.Create(Self);
  btn9.Name := 'btn9';
  btn9.Caption := '9';
  btn9.Left := 136;
  btn9.Top := 60;
  btn9.Width := 58;
  btn9.Height := 48;
  btn9.OnClick := Btn9Click;
  Self.Controls.Add(btn9);

  btnDiv := TButton.Create(Self);
  btnDiv.Name := 'btnDiv';
  btnDiv.Caption := '/';
  btnDiv.Left := 199;
  btnDiv.Top := 60;
  btnDiv.Width := 58;
  btnDiv.Height := 48;
  btnDiv.OnClick := BtnDivClick;
  Self.Controls.Add(btnDiv);

  { Row 2: 4 5 6 * }
  btn4 := TButton.Create(Self);
  btn4.Name := 'btn4';
  btn4.Caption := '4';
  btn4.Left := 10;
  btn4.Top := 115;
  btn4.Width := 58;
  btn4.Height := 48;
  btn4.OnClick := Btn4Click;
  Self.Controls.Add(btn4);

  btn5 := TButton.Create(Self);
  btn5.Name := 'btn5';
  btn5.Caption := '5';
  btn5.Left := 73;
  btn5.Top := 115;
  btn5.Width := 58;
  btn5.Height := 48;
  btn5.OnClick := Btn5Click;
  Self.Controls.Add(btn5);

  btn6 := TButton.Create(Self);
  btn6.Name := 'btn6';
  btn6.Caption := '6';
  btn6.Left := 136;
  btn6.Top := 115;
  btn6.Width := 58;
  btn6.Height := 48;
  btn6.OnClick := Btn6Click;
  Self.Controls.Add(btn6);

  btnMul := TButton.Create(Self);
  btnMul.Name := 'btnMul';
  btnMul.Caption := '*';
  btnMul.Left := 199;
  btnMul.Top := 115;
  btnMul.Width := 58;
  btnMul.Height := 48;
  btnMul.OnClick := BtnMulClick;
  Self.Controls.Add(btnMul);

  { Row 3: 1 2 3 - }
  btn1 := TButton.Create(Self);
  btn1.Name := 'btn1';
  btn1.Caption := '1';
  btn1.Left := 10;
  btn1.Top := 170;
  btn1.Width := 58;
  btn1.Height := 48;
  btn1.OnClick := Btn1Click;
  Self.Controls.Add(btn1);

  btn2 := TButton.Create(Self);
  btn2.Name := 'btn2';
  btn2.Caption := '2';
  btn2.Left := 73;
  btn2.Top := 170;
  btn2.Width := 58;
  btn2.Height := 48;
  btn2.OnClick := Btn2Click;
  Self.Controls.Add(btn2);

  btn3 := TButton.Create(Self);
  btn3.Name := 'btn3';
  btn3.Caption := '3';
  btn3.Left := 136;
  btn3.Top := 170;
  btn3.Width := 58;
  btn3.Height := 48;
  btn3.OnClick := Btn3Click;
  Self.Controls.Add(btn3);

  btnSub := TButton.Create(Self);
  btnSub.Name := 'btnSub';
  btnSub.Caption := '-';
  btnSub.Left := 199;
  btnSub.Top := 170;
  btnSub.Width := 58;
  btnSub.Height := 48;
  btnSub.OnClick := BtnSubClick;
  Self.Controls.Add(btnSub);

  { Row 4: C 0 = + }
  btnC := TButton.Create(Self);
  btnC.Name := 'btnC';
  btnC.Caption := 'C';
  btnC.Left := 10;
  btnC.Top := 225;
  btnC.Width := 58;
  btnC.Height := 48;
  btnC.OnClick := BtnCClick;
  Self.Controls.Add(btnC);

  btn0 := TButton.Create(Self);
  btn0.Name := 'btn0';
  btn0.Caption := '0';
  btn0.Left := 73;
  btn0.Top := 225;
  btn0.Width := 58;
  btn0.Height := 48;
  btn0.OnClick := Btn0Click;
  Self.Controls.Add(btn0);

  btnEq := TButton.Create(Self);
  btnEq.Name := 'btnEq';
  btnEq.Caption := '=';
  btnEq.Left := 136;
  btnEq.Top := 225;
  btnEq.Width := 58;
  btnEq.Height := 48;
  btnEq.OnClick := BtnEqClick;
  Self.Controls.Add(btnEq);

  btnAdd := TButton.Create(Self);
  btnAdd.Name := 'btnAdd';
  btnAdd.Caption := '+';
  btnAdd.Left := 199;
  btnAdd.Top := 225;
  btnAdd.Width := 58;
  btnAdd.Height := 48;
  btnAdd.OnClick := BtnAddClick;
  Self.Controls.Add(btnAdd);
end;

procedure TCalcForm.UpdateDisplay;
begin
  display.Text := FCurrent;
end;

procedure TCalcForm.PressDigit(d: string);
begin
  if FResetNext then
  begin
    FCurrent := d;
    FResetNext := False;
  end
  else
  begin
    if FCurrent = '0' then
      FCurrent := d
    else
      FCurrent := FCurrent + d;
  end;
  UpdateDisplay;
end;

procedure TCalcForm.PressOperator(op: string);
begin
  if (FPrevious <> '') and (not FResetNext) then
    DoCalculate;
  FPrevious := FCurrent;
  FOp := op;
  FResetNext := True;
end;

procedure TCalcForm.DoCalculate;
var
  a, b, result: Double;
begin
  if (FPrevious = '') or (FOp = '') then
    Exit;
  a := StrToFloat(FPrevious);
  b := StrToFloat(FCurrent);
  result := 0;
  if FOp = '+' then result := a + b;
  if FOp = '-' then result := a - b;
  if FOp = '*' then result := a * b;
  if FOp = '/' then
  begin
    if b = 0 then
    begin
      FCurrent := 'Error';
      FPrevious := '';
      FOp := '';
      FResetNext := True;
      UpdateDisplay;
      Exit;
    end;
    result := a / b;
  end;
  FCurrent := FloatToStr(result);
  FPrevious := '';
  FOp := '';
  FResetNext := True;
  UpdateDisplay;
end;

procedure TCalcForm.PressClear;
begin
  FCurrent := '0';
  FPrevious := '';
  FOp := '';
  FResetNext := False;
  UpdateDisplay;
end;

procedure TCalcForm.Btn7Click(Sender: TObject); begin PressDigit('7'); end;
procedure TCalcForm.Btn8Click(Sender: TObject); begin PressDigit('8'); end;
procedure TCalcForm.Btn9Click(Sender: TObject); begin PressDigit('9'); end;
procedure TCalcForm.BtnDivClick(Sender: TObject); begin PressOperator('/'); end;
procedure TCalcForm.Btn4Click(Sender: TObject); begin PressDigit('4'); end;
procedure TCalcForm.Btn5Click(Sender: TObject); begin PressDigit('5'); end;
procedure TCalcForm.Btn6Click(Sender: TObject); begin PressDigit('6'); end;
procedure TCalcForm.BtnMulClick(Sender: TObject); begin PressOperator('*'); end;
procedure TCalcForm.Btn1Click(Sender: TObject); begin PressDigit('1'); end;
procedure TCalcForm.Btn2Click(Sender: TObject); begin PressDigit('2'); end;
procedure TCalcForm.Btn3Click(Sender: TObject); begin PressDigit('3'); end;
procedure TCalcForm.BtnSubClick(Sender: TObject); begin PressOperator('-'); end;
procedure TCalcForm.BtnCClick(Sender: TObject); begin PressClear; end;
procedure TCalcForm.Btn0Click(Sender: TObject); begin PressDigit('0'); end;
procedure TCalcForm.BtnEqClick(Sender: TObject); begin DoCalculate; end;
procedure TCalcForm.BtnAddClick(Sender: TObject); begin PressOperator('+'); end;

var
  CalcForm: TCalcForm;
begin
  Application.Initialize;
  Application.CreateForm(TCalcForm, CalcForm);
  Application.Run;
end.
