// vybe-test: pascal/properties_extended/property_write_only_via_setter_side_effect
// origin: languages/pascal/tests/pascal/test_properties_extended.rs
program T;
{$mode delphi}
uses SysUtils; type TLog=class private FLast:Integer; procedure SetTarget(v:Integer); public property Target:Integer write SetTarget; function Last:Integer; begin Result:=FLast; end; end; procedure TLog.SetTarget(v:Integer); begin FLast:=v; end; var L:TLog; begin L:=TLog.Create; L.Target:=42; WriteLn(L.Last); end.
