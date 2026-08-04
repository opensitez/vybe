// vybe-test: pascal/file_text_extra/text_blank_line_roundtrips_as_empty_string
// origin: languages/pascal/tests/pascal/test_file_text_extra.rs
program T;
{$mode delphi}
uses SysUtils; var f: TextFile; s: string; begin Assign(f,'text_blank.txt'); Rewrite(f); WriteLn(f); WriteLn(f,'x'); Close(f); Reset(f); ReadLn(f,s); if s = '' then WriteLn('blank'); ReadLn(f,s); Close(f); WriteLn(s); end.
