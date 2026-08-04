// vybe-test: pascal/file_io_extended/erase_removes_named_file
// origin: languages/pascal/tests/pascal/test_file_io_extended.rs
program T;
{$mode delphi}
uses SysUtils; var f: TextFile; begin Assign(f,'ext_erase.txt'); Rewrite(f); WriteLn(f,'gone'); Close(f); Erase(f); if not FileExists('ext_erase.txt') then WriteLn('gone'); end.
