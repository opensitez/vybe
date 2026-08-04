// vybe-test: pascal/file_io/typed_file_multiple_records_preserve_order
// origin: languages/pascal/tests/pascal/test_file_io.rs
program T;
{$mode delphi}
uses SysUtils; type TIntFile = file of Integer; var f: TIntFile; a,b: Integer; begin Assign(f,'core_multi.dat'); Rewrite(f); a := 3; b := 4; Write(f,a); Write(f,b); Close(f); Reset(f); a := 0; b := 0; Read(f,a); Read(f,b); Close(f); WriteLn(a+b); end.
