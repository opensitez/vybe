program ColorMixerApp;

uses
  Forms,
  MainForm in 'MainForm.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.Title := 'VCL Color Mixer Studio';
  Application.CreateForm(TColorMixerForm, ColorMixerForm);
  Application.Run;
end.
