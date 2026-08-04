// vybe-test: pascal/properties_extended/property_setter_clamps_upper_bound
// origin: languages/pascal/tests/pascal/test_properties_extended.rs
program T;
{$mode delphi}
uses SysUtils; type TPercent=class private F:Integer; procedure SetP(v:Integer); public property P:Integer read F write SetP; end; procedure TPercent.SetP(v:Integer); begin if v>100 then F:=100 else F:=v; end; var p:TPercent; begin p:=TPercent.Create; p.P:=150; WriteLn(p.P); end.
