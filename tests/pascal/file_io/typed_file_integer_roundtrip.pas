// vybe-test: pascal/file_io/typed_file_integer_roundtrip
// origin: languages/pascal/tests/pascal/test_file_io.rs
program T;
{$mode delphi}
uses SysUtils; var f: file of Integer; n: Integer; begin Assign(f,'core_ints.dat'); Rewrite(f); n := 7; Write(f,n); Close(f); Reset(f); n := 0; Read(f,n); Close(f); WriteLn(n); end.
