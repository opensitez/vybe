// vybe-test: pascal/properties_extended/property_write_string_trim_on_set
// origin: languages/pascal/tests/pascal/test_properties_extended.rs
program T;
{$mode delphi}
uses SysUtils; type TTrim=class private FS:string; procedure SetS(const v:string); public property S:string read FS write SetS; end; procedure TTrim.SetS(const v:string); begin FS:=Trim(v); end; var t:TTrim; begin t:=TTrim.Create; t.S:='  ok  '; WriteLn(t.S); end.
