program DelphiDemo;

uses
  Forms,
  MainForm in 'MainForm.pas',
  SettingsForm in 'SettingsForm.pas',
  DataForm in 'DataForm.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'Delphi Multi-Form Demo';
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
