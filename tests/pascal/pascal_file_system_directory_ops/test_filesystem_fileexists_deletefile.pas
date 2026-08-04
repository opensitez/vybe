// vybe-test: pascal/pascal_file_system_directory_ops/test_filesystem_fileexists_deletefile
// origin: languages/pascal/tests/pascal/test_pascal_file_system_directory_ops.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: TextFile; fname: String;
begin
  fname := 'test_fs_del.tmp';
  AssignFile(f, fname); Rewrite(f); WriteLn(f, 'tmp'); CloseFile(f);

  WriteLn(FileExists(fname));
  DeleteFile(fname);
  WriteLn(FileExists(fname));
end.
