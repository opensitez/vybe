// vybe-test: pascal/properties_extended/property_dynamic_grow_on_setter
// origin: languages/pascal/tests/pascal/test_properties_extended.rs
program T;
{$mode delphi}
uses SysUtils; type TList=class private F:array of Integer; procedure SetAt(i,v:Integer); function GetAt(i:Integer):Integer; public property At[i:Integer]:Integer read GetAt write SetAt; end; procedure TList.SetAt(i,v:Integer); begin if Length(F)<=i then SetLength(F,i+1); F[i]:=v; end; function TList.GetAt(i:Integer):Integer; begin Result:=F[i]; end; var L:TList; begin L:=TList.Create; L.At[2]:=77; WriteLn(L.At[2]); end.
