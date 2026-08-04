// vybe-test: pascal/properties_extended/property_setter_clamps_lower_bound
// origin: languages/pascal/tests/pascal/test_properties_extended.rs
program T;
{$mode delphi}
uses SysUtils; type TPercent=class private F:Integer; procedure SetP(v:Integer); public property P:Integer read F write SetP; end; procedure TPercent.SetP(v:Integer); begin if v<0 then F:=0 else F:=v; end; var p:TPercent; begin p:=TPercent.Create; p.P:=-5; WriteLn(p.P); end.
