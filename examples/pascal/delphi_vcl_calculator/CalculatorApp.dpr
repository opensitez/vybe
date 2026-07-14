program CalculatorApp;

uses
  Forms,
  MainForm in 'MainForm.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.Title := 'VCL Calculator';
  Application.CreateForm(TCalculatorForm, CalculatorForm);
  Application.Run;
end.
