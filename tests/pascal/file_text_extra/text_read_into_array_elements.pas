// vybe-test: pascal/file_text_extra/text_read_into_array_elements
// origin: languages/pascal/tests/pascal/test_file_text_extra.rs
program T;
{$mode delphi}
uses SysUtils; var f: TextFile; lines: array[0..1] of string; begin Assign(f,'text_array.txt'); Rewrite(f); WriteLn(f,'first'); WriteLn(f,'second'); Close(f); Reset(f); ReadLn(f,lines[0]); ReadLn(f,lines[1]); Close(f); WriteLn(lines[1]); WriteLn(lines[0]); end.
