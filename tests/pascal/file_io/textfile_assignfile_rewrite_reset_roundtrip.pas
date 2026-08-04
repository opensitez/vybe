// vybe-test: pascal/file_io/textfile_assignfile_rewrite_reset_roundtrip
// origin: languages/pascal/tests/pascal/test_file_io.rs
program T;
{$mode delphi}
uses SysUtils; var f: TextFile; s: string; begin AssignFile(f,'core_assign.txt'); Rewrite(f); WriteLn(f,'alpha'); CloseFile(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end.
