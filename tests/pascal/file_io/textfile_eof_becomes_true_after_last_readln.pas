// vybe-test: pascal/file_io/textfile_eof_becomes_true_after_last_readln
// origin: languages/pascal/tests/pascal/test_file_io.rs
program T;
{$mode delphi}
uses SysUtils; var f: TextFile; s: string; begin Assign(f,'core_eof.txt'); Rewrite(f); WriteLn(f,'last'); Close(f); Reset(f); if not Eof(f) then WriteLn('before'); ReadLn(f,s); if Eof(f) then WriteLn('after'); Close(f); end.
