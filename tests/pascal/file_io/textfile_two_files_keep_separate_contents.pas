// vybe-test: pascal/file_io/textfile_two_files_keep_separate_contents
// origin: languages/pascal/tests/pascal/test_file_io.rs
program T;
{$mode delphi}
uses SysUtils; var a, b: TextFile; s: string; begin Assign(a,'core_a.txt'); Rewrite(a); WriteLn(a,'A'); Close(a); Assign(b,'core_b.txt'); Rewrite(b); WriteLn(b,'B'); Close(b); Reset(a); ReadLn(a,s); Close(a); WriteLn(s); Reset(b); ReadLn(b,s); Close(b); WriteLn(s); end.
