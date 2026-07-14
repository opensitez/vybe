program KanbanLiteApp;

uses
  Forms,
  MainForm in 'MainForm.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.Title := 'VCL Kanban Lite';
  Application.CreateForm(TKanbanForm, KanbanForm);
  Application.Run;
end.
