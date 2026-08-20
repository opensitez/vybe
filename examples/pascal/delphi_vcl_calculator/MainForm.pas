unit MainForm;

interface

uses
  SysUtils, Classes, Forms, Controls, StdCtrls, ExtCtrls, Dialogs;

type
  TCalculatorForm = class(TForm)
  private
    FDisplay: TEdit;
    FPendingOp: Char;
    FStoredValue: Double;
    FResetOnNextDigit: Boolean;
    procedure BuildUi;
    procedure NumberClick(Sender: TObject);
    procedure DecimalClick(Sender: TObject);
    procedure OperatorClick(Sender: TObject);
    procedure EqualsClick(Sender: TObject);
    procedure ClearClick(Sender: TObject);
    procedure BackspaceClick(Sender: TObject);
    procedure ApplyPendingOperation(const NewValue: Double);
    function DisplayValue: Double;
    procedure SetDisplayValue(const Value: Double);
  public
    constructor Create(AOwner: TComponent); override;
  end;

var
  CalculatorForm: TCalculatorForm;

implementation

{$R *.dfm}

constructor TCalculatorForm.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  BuildUi;
  FPendingOp := #0;
  FStoredValue := 0;
  FResetOnNextDigit := False;
end;

procedure TCalculatorForm.BuildUi;
const
  Labels: array[0..15] of string = (
    '7', '8', '9', '/',
    '4', '5', '6', '*',
    '1', '2', '3', '-',
    '0', '.', '=', '+'
  );
var
  I: Integer;
  Btn: TButton;
  Col, Row: Integer;
  TopPanel: TPanel;
begin
  Caption := 'Calculator';
  Width := 330;
  Height := 410;
  Position := poScreenCenter;

  TopPanel := TPanel.Create(Self);
  TopPanel.Parent := Self;
  TopPanel.Align := alTop;
  TopPanel.Height := 72;
  TopPanel.BevelOuter := bvNone;

  FDisplay := TEdit.Create(Self);
  FDisplay.Parent := TopPanel;
  FDisplay.Left := 8;
  FDisplay.Top := 8;
  FDisplay.Width := 296;
  FDisplay.Height := 42;
  FDisplay.Alignment := taRightJustify;
  FDisplay.Font.Size := 16;
  FDisplay.ReadOnly := True;
  FDisplay.Text := '0';

  Btn := TButton.Create(Self);
  Btn.Parent := TopPanel;
  Btn.Caption := 'C';
  Btn.Left := 8;
  Btn.Top := 50;
  Btn.Width := 146;
  Btn.Height := 20;
  Btn.OnClick := ClearClick;

  Btn := TButton.Create(Self);
  Btn.Parent := TopPanel;
  Btn.Caption := 'Back';
  Btn.Left := 158;
  Btn.Top := 50;
  Btn.Width := 146;
  Btn.Height := 20;
  Btn.OnClick := BackspaceClick;

  for I := 0 to High(Labels) do
  begin
    Btn := TButton.Create(Self);
    Btn.Parent := Self;
    Col := I mod 4;
    Row := I div 4;
    Btn.Left := 8 + (Col * 74);
    Btn.Top := 88 + (Row * 64);
    Btn.Width := 70;
    Btn.Height := 56;
    Btn.Caption := Labels[I];

    if (Labels[I][1] in ['0'..'9']) then
      Btn.OnClick := NumberClick
    else if Labels[I] = '.' then
      Btn.OnClick := DecimalClick
    else if Labels[I] = '=' then
      Btn.OnClick := EqualsClick
    else
      Btn.OnClick := OperatorClick;
  end;
end;

procedure TCalculatorForm.NumberClick(Sender: TObject);
var
  Digit: string;
begin
  Digit := (Sender as TButton).Caption;

  if FResetOnNextDigit then
  begin
    FDisplay.Text := '0';
    FResetOnNextDigit := False;
  end;

  if FDisplay.Text = '0' then
    FDisplay.Text := Digit
  else
    FDisplay.Text := FDisplay.Text + Digit;
end;

procedure TCalculatorForm.DecimalClick(Sender: TObject);
begin
  if FResetOnNextDigit then
  begin
    FDisplay.Text := '0';
    FResetOnNextDigit := False;
  end;

  if Pos('.', FDisplay.Text) = 0 then
    FDisplay.Text := FDisplay.Text + '.';
end;

procedure TCalculatorForm.OperatorClick(Sender: TObject);
begin
  ApplyPendingOperation(DisplayValue);
  FPendingOp := (Sender as TButton).Caption[1];
  FResetOnNextDigit := True;
end;

procedure TCalculatorForm.EqualsClick(Sender: TObject);
begin
  ApplyPendingOperation(DisplayValue);
  FPendingOp := #0;
  SetDisplayValue(FStoredValue);
  FResetOnNextDigit := True;
end;

procedure TCalculatorForm.ClearClick(Sender: TObject);
begin
  FDisplay.Text := '0';
  FStoredValue := 0;
  FPendingOp := #0;
  FResetOnNextDigit := False;
end;

procedure TCalculatorForm.BackspaceClick(Sender: TObject);
var
  S: string;
begin
  if FResetOnNextDigit then
    Exit;

  { `Delete` takes a `var string`, and `Text` is a property — it cannot be
    passed by reference. Edit a local and assign the result back. }
  S := FDisplay.Text;
  if Length(S) <= 1 then
    FDisplay.Text := '0'
  else
  begin
    Delete(S, Length(S), 1);
    FDisplay.Text := S;
  end;
end;

procedure TCalculatorForm.ApplyPendingOperation(const NewValue: Double);
begin
  if FPendingOp = #0 then
  begin
    FStoredValue := NewValue;
    Exit;
  end;

  case FPendingOp of
    '+': FStoredValue := FStoredValue + NewValue;
    '-': FStoredValue := FStoredValue - NewValue;
    '*': FStoredValue := FStoredValue * NewValue;
    '/':
      begin
        if NewValue = 0 then
        begin
          ShowMessage('Cannot divide by zero.');
          FPendingOp := #0;
          Exit;
        end;
        FStoredValue := FStoredValue / NewValue;
      end;
  end;

  SetDisplayValue(FStoredValue);
end;

function TCalculatorForm.DisplayValue: Double;
begin
  Result := StrToFloatDef(FDisplay.Text, 0);
end;

procedure TCalculatorForm.SetDisplayValue(const Value: Double);
begin
  FDisplay.Text := FloatToStr(Value);
end;

end.
