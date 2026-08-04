// vybe-test: pascal/file_text_extra/text_writeln_no_arguments_writes_empty_line
// origin: languages/pascal/tests/pascal/test_file_text_extra.rs
program T;
{$mode delphi}
uses SysUtils; var f: TextFile; s: string; begin Assign(f,'text_empty_line.txt'); Rewrite(f); WriteLn(f); Close(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(Length(s)); end.
