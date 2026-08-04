// vybe-test: pascal/file_io_extended/typed_file_write_multiple_arguments_appends_records
// origin: languages/pascal/tests/pascal/test_file_io_extended.rs
program T;
{$mode delphi}
uses SysUtils; type TIntFile = file of Integer; var f: TIntFile; a,b,c: Integer; begin Assign(f,'ext_multi_args.dat'); Rewrite(f); a := 1; b := 2; c := 3; Write(f,a,b,c); Close(f); Reset(f); a := 0; b := 0; c := 0; Read(f,a,b,c); Close(f); WriteLn(a+b+c); end.
