// vybe-test: pascal/properties_extended/property_dynamic_array_indexed_write_read
// origin: languages/pascal/tests/pascal/test_properties_extended.rs
program T;
{$mode delphi}
uses SysUtils; type TBuf=class private F:array of string; public property Cells[i:Integer]:string read GetCell write SetCell; function GetCell(i:Integer):string; begin Result:=F[i]; end; procedure SetCell(i:Integer; const v:string); begin F[i]:=v; end; procedure Grow(n:Integer); begin SetLength(F,n); end; end; var b:TBuf; begin b:=TBuf.Create; b.Grow(2); b.Cells[0]:='a'; b.Cells[1]:='b'; WriteLn(b.Cells[1]); end.
