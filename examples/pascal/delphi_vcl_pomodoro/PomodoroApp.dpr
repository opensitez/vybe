program PomodoroApp;

uses
  Forms,
  MainForm in 'MainForm.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.Title := 'VCL Pomodoro Focus Timer';
  Application.CreateForm(TPomodoroForm, PomodoroForm);
  Application.Run;
end.
