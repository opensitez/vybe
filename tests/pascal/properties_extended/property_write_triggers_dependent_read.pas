// vybe-test: pascal/properties_extended/property_write_triggers_dependent_read
// origin: languages/pascal/tests/pascal/test_properties_extended.rs
program T;
{$mode delphi}
uses SysUtils; type TC=class private FR,FC:Integer; procedure SetR(v:Integer); function GetC:Integer; public property R:Integer read FR write SetR; property C:Integer read GetC; end; procedure TC.SetR(v:Integer); begin FR:=v; FC:=v*2; end; function TC.GetC:Integer; begin Result:=FC; end; var c:TC; begin c:=TC.Create; c.R:=5; WriteLn(c.C); end.
