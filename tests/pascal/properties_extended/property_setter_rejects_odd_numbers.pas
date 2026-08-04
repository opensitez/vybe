// vybe-test: pascal/properties_extended/property_setter_rejects_odd_numbers
// origin: languages/pascal/tests/pascal/test_properties_extended.rs
program T;
{$mode delphi}
uses SysUtils; type TEven=class private F:Integer; procedure SetE(v:Integer); public property Even:Integer read F write SetE; end; procedure TEven.SetE(v:Integer); begin if v mod 2=0 then F:=v; end; var e:TEven; begin e:=TEven.Create; e.Even:=3; WriteLn(e.Even); e.Even:=4; WriteLn(e.Even); end.
