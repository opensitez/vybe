unit MainForm;

interface

uses
  SysUtils, Classes, Forms, Controls, StdCtrls, ExtCtrls;

type
  TPomodoroForm = class(TForm)
  private
    FTimer: TTimer;
    FRemainingSeconds: Integer;
    FRunning: Boolean;
    FTimeLabel: TLabel;
    FStateLabel: TLabel;
    FModeLabel: TLabel;
    FStartPauseButton: TButton;
    FResetButton: TButton;
    FWorkButton: TButton;
    FBreakButton: TButton;
    procedure BuildUi;
    procedure SetTimerDuration(Minutes: Integer);
    procedure Tick(Sender: TObject);
    procedure StartPauseClick(Sender: TObject);
    procedure ResetClick(Sender: TObject);
    procedure WorkClick(Sender: TObject);
    procedure BreakClick(Sender: TObject);
    procedure UpdateDisplay;
  public
    constructor Create(AOwner: TComponent); override;
  end;

var
  PomodoroForm: TPomodoroForm;

implementation

{$R *.dfm}

constructor TPomodoroForm.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  BuildUi;

  FTimer := TTimer.Create(Self);
  FTimer.Interval := 1000;
  FTimer.Enabled := False;
  FTimer.OnTimer := Tick;

  SetTimerDuration(25);
end;

procedure TPomodoroForm.BuildUi;
begin
  Caption := 'Pomodoro Focus Timer';
  Width := 420;
  Height := 280;
  Position := poScreenCenter;

  FTimeLabel := TLabel.Create(Self);
  FTimeLabel.Parent := Self;
  FTimeLabel.Left := 30;
  FTimeLabel.Top := 24;
  FTimeLabel.Font.Size := 36;
  FTimeLabel.Font.Style := [fsBold];

  FStateLabel := TLabel.Create(Self);
  FStateLabel.Parent := Self;
  FStateLabel.Left := 30;
  FStateLabel.Top := 90;
  FStateLabel.Font.Size := 11;

  FModeLabel := TLabel.Create(Self);
  FModeLabel.Parent := Self;
  FModeLabel.Left := 30;
  FModeLabel.Top := 110;
  FModeLabel.Font.Size := 10;
  FModeLabel.Caption := 'Mode: Focus session';

  FStartPauseButton := TButton.Create(Self);
  FStartPauseButton.Parent := Self;
  FStartPauseButton.Left := 30;
  FStartPauseButton.Top := 130;
  FStartPauseButton.Width := 110;
  FStartPauseButton.Height := 34;
  FStartPauseButton.Caption := 'Start';
  FStartPauseButton.OnClick := StartPauseClick;

  FResetButton := TButton.Create(Self);
  FResetButton.Parent := Self;
  FResetButton.Left := 150;
  FResetButton.Top := 130;
  FResetButton.Width := 110;
  FResetButton.Height := 34;
  FResetButton.Caption := 'Reset';
  FResetButton.OnClick := ResetClick;

  FWorkButton := TButton.Create(Self);
  FWorkButton.Parent := Self;
  FWorkButton.Left := 30;
  FWorkButton.Top := 180;
  FWorkButton.Width := 110;
  FWorkButton.Height := 30;
  FWorkButton.Caption := 'Work 25m';
  FWorkButton.OnClick := WorkClick;

  FBreakButton := TButton.Create(Self);
  FBreakButton.Parent := Self;
  FBreakButton.Left := 150;
  FBreakButton.Top := 180;
  FBreakButton.Width := 110;
  FBreakButton.Height := 30;
  FBreakButton.Caption := 'Break 5m';
  FBreakButton.OnClick := BreakClick;
end;

procedure TPomodoroForm.SetTimerDuration(Minutes: Integer);
begin
  FRemainingSeconds := Minutes * 60;
  FRunning := False;
  FTimer.Enabled := False;
  FStartPauseButton.Caption := 'Start';
  UpdateDisplay;
end;

procedure TPomodoroForm.Tick(Sender: TObject);
begin
  if FRemainingSeconds > 0 then
    Dec(FRemainingSeconds)
  else
  begin
    FTimer.Enabled := False;
    FRunning := False;
    FStartPauseButton.Caption := 'Start';
    ShowMessage('Time is up. Great work!');
  end;

  UpdateDisplay;
end;

procedure TPomodoroForm.StartPauseClick(Sender: TObject);
begin
  FRunning := not FRunning;
  FTimer.Enabled := FRunning;

  if FRunning then
    FStartPauseButton.Caption := 'Pause'
  else
    FStartPauseButton.Caption := 'Start';

  UpdateDisplay;
end;

procedure TPomodoroForm.ResetClick(Sender: TObject);
begin
  SetTimerDuration(25);
end;

procedure TPomodoroForm.WorkClick(Sender: TObject);
begin
  SetTimerDuration(25);
  FModeLabel.Caption := 'Mode: Focus session';
end;

procedure TPomodoroForm.BreakClick(Sender: TObject);
begin
  SetTimerDuration(5);
  FModeLabel.Caption := 'Mode: Break session';
end;

procedure TPomodoroForm.UpdateDisplay;
var
  MinPart, SecPart: Integer;
begin
  MinPart := FRemainingSeconds div 60;
  SecPart := FRemainingSeconds mod 60;
  FTimeLabel.Caption := Format('%.2d:%.2d', [MinPart, SecPart]);

  if FRunning then
    FStateLabel.Caption := 'Status: Running'
  else if FRemainingSeconds = 0 then
    FStateLabel.Caption := 'Status: Completed'
  else
    FStateLabel.Caption := 'Status: Paused';
end;

end.
