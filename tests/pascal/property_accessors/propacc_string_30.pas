// vybe-test: pascal/property_accessors/propacc_string_30
// origin: languages/pascal/tests/pascal/test_property_accessors.rs
program T;
{$mode delphi}
uses SysUtils; type T=class private FName:string; function GetName:string; procedure SetName(const s:string); public property Name:string read GetName write SetName; end; function T.GetName:string; begin Result:=FName; end; procedure T.SetName(const s:string); begin FName:=s+'_30'; end; var o:T; begin o:=T.Create; o.Name:='x'; WriteLn(o.Name); o.Free; end.
