// vybe-test: pascal/property_accessors/propacc_rw_custom_29
// origin: languages/pascal/tests/pascal/test_property_accessors.rs
program T;
{$mode delphi}
uses SysUtils; type T=class private F:Integer; function GetD:Integer; procedure SetD(v:Integer); public property Double:Integer read GetD write SetD; end; function T.GetD:Integer; begin Result:=F*2; end; procedure T.SetD(v:Integer); begin F:=v div 2; end; var o:T; begin o:=T.Create; o.Double:=58; WriteLn(o.Double); o.Free; end.
