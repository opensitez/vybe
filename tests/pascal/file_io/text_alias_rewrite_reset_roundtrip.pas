// vybe-test: pascal/file_io/text_alias_rewrite_reset_roundtrip
// origin: languages/pascal/tests/pascal/test_file_io.rs
program T;
{$mode delphi}
uses SysUtils; var f: Text; s: string; begin Assign(f,'core_text_alias.txt'); Rewrite(f); WriteLn(f,'alias'); Close(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end.
