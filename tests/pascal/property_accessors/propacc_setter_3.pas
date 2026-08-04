// vybe-test: pascal/property_accessors/propacc_setter_3
// origin: languages/pascal/tests/pascal/test_property_accessors.rs
program T;
{$mode delphi}
uses SysUtils; type T=class private F:Integer; procedure SetV(v:Integer); public property Val:Integer read F write SetV; end; procedure T.SetV(v:Integer); begin F:=v+3; end; var o:T; begin o:=T.Create; o.Val:=3; WriteLn(o.F); o.Free; end.
