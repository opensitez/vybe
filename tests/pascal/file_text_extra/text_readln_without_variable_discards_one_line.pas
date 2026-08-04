// vybe-test: pascal/file_text_extra/text_readln_without_variable_discards_one_line
// origin: languages/pascal/tests/pascal/test_file_text_extra.rs
program T;
{$mode delphi}
uses SysUtils; var f: TextFile; s: string; begin Assign(f,'text_discard.txt'); Rewrite(f); WriteLn(f,'skip'); WriteLn(f,'keep'); Close(f); Reset(f); ReadLn(f); ReadLn(f,s); Close(f); WriteLn(s); end.
