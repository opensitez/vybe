// vybe-test: pascal/file_text_extra/text_long_line_roundtrips_without_truncation
// origin: languages/pascal/tests/pascal/test_file_text_extra.rs
program T;
{$mode delphi}
uses SysUtils; var f: TextFile; s: string; begin Assign(f,'text_long.txt'); Rewrite(f); WriteLn(f,'abcdefghijklmnopqrstuvwxyz'); Close(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(Length(s)); end.
