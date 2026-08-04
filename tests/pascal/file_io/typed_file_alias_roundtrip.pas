// vybe-test: pascal/file_io/typed_file_alias_roundtrip
// origin: languages/pascal/tests/pascal/test_file_io.rs
program T;
{$mode delphi}
uses SysUtils; type TIntFile = file of Integer; var f: TIntFile; n: Integer; begin Assign(f,'core_alias.dat'); Rewrite(f); n := 42; Write(f,n); Close(f); Reset(f); n := 0; Read(f,n); Close(f); WriteLn(n); end.
