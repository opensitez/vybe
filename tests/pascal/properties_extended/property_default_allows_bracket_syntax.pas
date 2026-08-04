// vybe-test: pascal/properties_extended/property_default_allows_bracket_syntax
// origin: languages/pascal/tests/pascal/test_properties_extended.rs
program T;
{$mode delphi}
uses SysUtils; type TMap=class private F:array[0..1] of Integer; public property Items[i:Integer]:Integer read GetItem write SetItem; default; function GetItem(i:Integer):Integer; begin Result:=F[i]; end; procedure SetItem(i:Integer; v:Integer); begin F[i]:=v; end; end; var m:TMap; begin m:=TMap.Create; m[0]:=9; m[1]:=4; WriteLn(m[1]); end.
