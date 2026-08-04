// vybe-test: pascal/file_io/textfile_rewrite_truncates_existing_content
// origin: languages/pascal/tests/pascal/test_file_io.rs
program T;
{$mode delphi}
uses SysUtils; var f: TextFile; s: string; begin Assign(f,'core_trunc.txt'); Rewrite(f); WriteLn(f,'old'); Close(f); Rewrite(f); WriteLn(f,'new'); Close(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end.
