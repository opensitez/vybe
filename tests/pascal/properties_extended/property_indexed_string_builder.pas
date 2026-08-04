// vybe-test: pascal/properties_extended/property_indexed_string_builder
// origin: languages/pascal/tests/pascal/test_properties_extended.rs
program T;
{$mode delphi}
uses SysUtils; type TParts=class private F:array[0..1] of string; function GetPart(i:Integer):string; procedure SetPart(i:Integer; const s:string); public property Part[i:Integer]:string read GetPart write SetPart; function Join:string; var r:string; begin r:=Part[0]+Part[1]; Result:=r; end; end; function TParts.GetPart(i:Integer):string; begin Result:=F[i]; end; procedure TParts.SetPart(i:Integer; const s:string); begin F[i]:=s; end; var p:TParts; begin p:=TParts.Create; p.Part[0]:='ab'; p.Part[1]:='cd'; WriteLn(p.Join); end.
