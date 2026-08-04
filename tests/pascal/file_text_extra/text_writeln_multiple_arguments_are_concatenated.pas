// vybe-test: pascal/file_text_extra/text_writeln_multiple_arguments_are_concatenated
// origin: languages/pascal/tests/pascal/test_file_text_extra.rs
program T;
{$mode delphi}
uses SysUtils; var f: TextFile; s: string; begin Assign(f,'text_multi_write.txt'); Rewrite(f); WriteLn(f,'A',1,'B',2); Close(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end.
