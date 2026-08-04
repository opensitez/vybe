// vybe-test: pascal/file_io/typed_file_eof_after_last_record
// origin: languages/pascal/tests/pascal/test_file_io.rs
program T;
{$mode delphi}
uses SysUtils; type TIntFile = file of Integer; var f: TIntFile; n: Integer; begin Assign(f,'core_typed_eof.dat'); Rewrite(f); n := 1; Write(f,n); Close(f); Reset(f); Read(f,n); if Eof(f) then WriteLn('eof'); Close(f); end.
