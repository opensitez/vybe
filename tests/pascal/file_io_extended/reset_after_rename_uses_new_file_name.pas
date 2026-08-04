// vybe-test: pascal/file_io_extended/reset_after_rename_uses_new_file_name
// origin: languages/pascal/tests/pascal/test_file_io_extended.rs
program T;
{$mode delphi}
uses SysUtils; var f: TextFile; s: string; begin Assign(f,'ext_rn1.txt'); Rewrite(f); WriteLn(f,'rn'); Close(f); Rename(f,'ext_rn2.txt'); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end.
