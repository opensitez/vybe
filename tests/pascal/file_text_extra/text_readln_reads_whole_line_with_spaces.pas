// vybe-test: pascal/file_text_extra/text_readln_reads_whole_line_with_spaces
// origin: languages/pascal/tests/pascal/test_file_text_extra.rs
program T;
{$mode delphi}
uses SysUtils; var f: TextFile; s: string; begin Assign(f,'text_line.txt'); Rewrite(f); WriteLn(f,'one two three'); Close(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end.
