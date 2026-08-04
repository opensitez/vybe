// vybe-test: pascal/file_io/textfile_close_is_idempotent_for_reopen
// origin: languages/pascal/tests/pascal/test_file_io.rs
program T;
{$mode delphi}
uses SysUtils; var f: TextFile; s: string; begin Assign(f,'core_close.txt'); Rewrite(f); WriteLn(f,'ok'); Close(f); CloseFile(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end.
