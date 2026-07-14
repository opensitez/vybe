program TextEditorApp;

uses
  Forms,
  MainForm in 'MainForm.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.Title := 'VCL Text Editor';
  Application.CreateForm(TTextEditorForm, TextEditorForm);
  Application.Run;
end.
