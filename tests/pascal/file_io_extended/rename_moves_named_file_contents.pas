// vybe-test: pascal/file_io_extended/rename_moves_named_file_contents
// origin: languages/pascal/tests/pascal/test_file_io_extended.rs
program T;
{$mode delphi}
uses SysUtils; var f: TextFile; s: string; begin Assign(f,'ext_old.txt'); Rewrite(f); WriteLn(f,'moved'); Close(f); Rename(f,'ext_new.txt'); if not FileExists('ext_old.txt') then WriteLn('old gone'); Assign(f,'ext_new.txt'); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end.
