// vybe-test: pascal/file_io_extended/typed_file_boolean_roundtrip
// origin: languages/pascal/tests/pascal/test_file_io_extended.rs
program T;
{$mode delphi}
uses SysUtils; type TBoolFile = file of Boolean; var f: TBoolFile; b: Boolean; begin Assign(f,'ext_bool.dat'); Rewrite(f); b := True; Write(f,b); Close(f); Reset(f); b := False; Read(f,b); Close(f); if b then WriteLn('true'); end.
