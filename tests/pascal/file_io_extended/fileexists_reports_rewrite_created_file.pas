// vybe-test: pascal/file_io_extended/fileexists_reports_rewrite_created_file
// origin: languages/pascal/tests/pascal/test_file_io_extended.rs
program T;
{$mode delphi}
uses SysUtils; var f: TextFile; begin Assign(f,'ext_exists.txt'); Rewrite(f); Close(f); if FileExists('ext_exists.txt') then WriteLn('exists'); end.
