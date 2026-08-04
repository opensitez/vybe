// vybe-test: pascal/properties_extended/property_notification_counter_on_write
// origin: languages/pascal/tests/pascal/test_properties_extended.rs
program T;
{$mode delphi}
uses SysUtils; type TWatch=class private FV,FChanges:Integer; procedure SetV(v:Integer); public property V:Integer read FV write SetV; property Changes:Integer read FChanges; end; procedure TWatch.SetV(v:Integer); begin FV:=v; Inc(FChanges); end; var w:TWatch; begin w:=TWatch.Create; w.V:=1; w.V:=2; WriteLn(w.Changes); end.
