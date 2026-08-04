// vybe-test: pascal/file_text_extra/text_boolean_token_reads_boolean_target
// origin: languages/pascal/tests/pascal/test_file_text_extra.rs
program T;
{$mode delphi}
uses SysUtils; var f: TextFile; b: Boolean; begin Assign(f,'text_bool.txt'); Rewrite(f); WriteLn(f,'True'); Close(f); Reset(f); ReadLn(f,b); Close(f); if b then WriteLn('true'); end.
