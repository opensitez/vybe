unit MainForm;

interface

uses
  SysUtils, Classes, Forms, Controls, StdCtrls;

type
  TTicTacToeForm = class(TForm)
  private
    FButtons: array[0..2, 0..2] of TButton;
    FCurrentPlayer: Char;
    FStatusLabel: TLabel;
    FScoreLabel: TLabel;
    FResetButton: TButton;
    FXWins: Integer;
    FOWins: Integer;
    procedure BuildUi;
    procedure CellClick(Sender: TObject);
    procedure ResetBoard(KeepScore: Boolean);
    procedure UpdateStatus;
    function CheckWinner(out Winner: Char): Boolean;
    function IsDraw: Boolean;
    procedure ResetButtonClick(Sender: TObject);
  public
    constructor Create(AOwner: TComponent); override;
  end;

var
  TicTacToeForm: TTicTacToeForm;

implementation

{$R *.dfm}

constructor TTicTacToeForm.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  BuildUi;
  FXWins := 0;
  FOWins := 0;
  ResetBoard(True);
end;

procedure TTicTacToeForm.BuildUi;
var
  Row, Col: Integer;
  Btn: TButton;
begin
  Caption := 'Tic Tac Toe';
  Width := 360;
  Height := 460;
  Position := poScreenCenter;

  FStatusLabel := TLabel.Create(Self);
  FStatusLabel.Parent := Self;
  FStatusLabel.Left := 20;
  FStatusLabel.Top := 16;
  FStatusLabel.Font.Size := 12;

  FScoreLabel := TLabel.Create(Self);
  FScoreLabel.Parent := Self;
  FScoreLabel.Left := 20;
  FScoreLabel.Top := 40;
  FScoreLabel.Font.Size := 10;

  for Row := 0 to 2 do
  begin
    for Col := 0 to 2 do
    begin
      Btn := TButton.Create(Self);
      Btn.Parent := Self;
      Btn.Left := 20 + (Col * 105);
      Btn.Top := 80 + (Row * 105);
      Btn.Width := 96;
      Btn.Height := 96;
      Btn.Font.Size := 24;
      Btn.Tag := Row * 3 + Col;
      Btn.OnClick := CellClick;
      FButtons[Row, Col] := Btn;
    end;
  end;

  FResetButton := TButton.Create(Self);
  FResetButton.Parent := Self;
  FResetButton.Left := 20;
  FResetButton.Top := 400;
  FResetButton.Width := 305;
  FResetButton.Height := 32;
  FResetButton.Caption := 'Reset Game';
  FResetButton.OnClick := ResetButtonClick;
end;

procedure TTicTacToeForm.CellClick(Sender: TObject);
var
  Btn: TButton;
  Winner: Char;
begin
  Btn := Sender as TButton;

  if Btn.Caption <> '' then
    Exit;

  Btn.Caption := FCurrentPlayer;

  if CheckWinner(Winner) then
  begin
    if Winner = 'X' then
      Inc(FXWins)
    else
      Inc(FOWins);

    ShowMessage(Format('Player %s wins!', [Winner]));
    ResetBoard(True);
    Exit;
  end;

  if IsDraw then
  begin
    ShowMessage('Draw game.');
    ResetBoard(True);
    Exit;
  end;

  if FCurrentPlayer = 'X' then
    FCurrentPlayer := 'O'
  else
    FCurrentPlayer := 'X';

  UpdateStatus;
end;

procedure TTicTacToeForm.ResetBoard(KeepScore: Boolean);
var
  Row, Col: Integer;
begin
  for Row := 0 to 2 do
    for Col := 0 to 2 do
      FButtons[Row, Col].Caption := '';

  if not KeepScore then
  begin
    FXWins := 0;
    FOWins := 0;
  end;

  FCurrentPlayer := 'X';
  UpdateStatus;
end;

procedure TTicTacToeForm.UpdateStatus;
begin
  FStatusLabel.Caption := Format('Current Player: %s', [FCurrentPlayer]);
  FScoreLabel.Caption := Format('Score  X: %d   O: %d', [FXWins, FOWins]);
end;

function TTicTacToeForm.CheckWinner(out Winner: Char): Boolean;
var
  I: Integer;
  Cells: array[0..2, 0..2] of string;
begin
  for I := 0 to 2 do
  begin
    Cells[I, 0] := FButtons[I, 0].Caption;
    Cells[I, 1] := FButtons[I, 1].Caption;
    Cells[I, 2] := FButtons[I, 2].Caption;

    if (Cells[I, 0] <> '') and (Cells[I, 0] = Cells[I, 1]) and (Cells[I, 1] = Cells[I, 2]) then
    begin
      Winner := Cells[I, 0][1];
      Exit(True);
    end;

    Cells[0, I] := FButtons[0, I].Caption;
    Cells[1, I] := FButtons[1, I].Caption;
    Cells[2, I] := FButtons[2, I].Caption;

    if (Cells[0, I] <> '') and (Cells[0, I] = Cells[1, I]) and (Cells[1, I] = Cells[2, I]) then
    begin
      Winner := Cells[0, I][1];
      Exit(True);
    end;
  end;

  if (FButtons[0, 0].Caption <> '') and
     (FButtons[0, 0].Caption = FButtons[1, 1].Caption) and
     (FButtons[1, 1].Caption = FButtons[2, 2].Caption) then
  begin
    Winner := FButtons[0, 0].Caption[1];
    Exit(True);
  end;

  if (FButtons[0, 2].Caption <> '') and
     (FButtons[0, 2].Caption = FButtons[1, 1].Caption) and
     (FButtons[1, 1].Caption = FButtons[2, 0].Caption) then
  begin
    Winner := FButtons[0, 2].Caption[1];
    Exit(True);
  end;

  Winner := #0;
  Result := False;
end;

function TTicTacToeForm.IsDraw: Boolean;
var
  Row, Col: Integer;
begin
  for Row := 0 to 2 do
    for Col := 0 to 2 do
      if FButtons[Row, Col].Caption = '' then
        Exit(False);

  Result := True;
end;

procedure TTicTacToeForm.ResetButtonClick(Sender: TObject);
begin
  ResetBoard(False);
end;

end.
