// vybe-test: pascal/properties_extended/property_indexed_write_then_sum
// origin: languages/pascal/tests/pascal/test_properties_extended.rs
program T;
{$mode delphi}
uses SysUtils; type TAcc=class private F:array[0..2] of Integer; function GetA(i:Integer):Integer; procedure SetA(i:Integer; v:Integer); public property A[i:Integer]:Integer read GetA write SetA; function Sum:Integer; var i,s:Integer; begin s:=0; for i:=0 to 2 do s:=s+A[i]; Result:=s; end; end; function TAcc.GetA(i:Integer):Integer; begin Result:=F[i]; end; procedure TAcc.SetA(i:Integer; v:Integer); begin F[i]:=v; end; var a:TAcc; begin a:=TAcc.Create; a.A[0]:=1; a.A[1]:=2; a.A[2]:=3; WriteLn(a.Sum); end.
