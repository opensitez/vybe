// vybe-test: pascal/properties_extended/property_array_of_char_indexed
// origin: languages/pascal/tests/pascal/test_properties_extended.rs
program T;
{$mode delphi}
uses SysUtils; type TChars=class private F:array[0..2] of Char; function GetCh(i:Integer):Char; procedure SetCh(i:Integer; c:Char); public property Ch[i:Integer]:Char read GetCh write SetCh; end; function TChars.GetCh(i:Integer):Char; begin Result:=F[i]; end; procedure TChars.SetCh(i:Integer; c:Char); begin F[i]:=c; end; var c:TChars; begin c:=TChars.Create; c.Ch[0]:='A'; c.Ch[1]:='B'; WriteLn(c.Ch[0]); WriteLn(c.Ch[1]); end.
