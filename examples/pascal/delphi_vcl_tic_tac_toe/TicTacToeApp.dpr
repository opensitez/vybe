program TicTacToeApp;

uses
  Forms,
  MainForm in 'MainForm.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.Title := 'VCL Tic Tac Toe';
  Application.CreateForm(TTicTacToeForm, TicTacToeForm);
  Application.Run;
end.
