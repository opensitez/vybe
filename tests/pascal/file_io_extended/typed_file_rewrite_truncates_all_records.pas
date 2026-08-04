// vybe-test: pascal/file_io_extended/typed_file_rewrite_truncates_all_records
// origin: languages/pascal/tests/pascal/test_file_io_extended.rs
program T;
{$mode delphi}
uses SysUtils; type TIntFile = file of Integer; var f: TIntFile; n: Integer; begin Assign(f,'ext_trunc.dat'); Rewrite(f); n := 1; Write(f,n); n := 2; Write(f,n); Close(f); Rewrite(f); n := 9; Write(f,n); Close(f); Reset(f); n := 0; Read(f,n); Close(f); WriteLn(n); end.
