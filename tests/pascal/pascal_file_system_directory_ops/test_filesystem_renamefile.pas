// vybe-test: pascal/pascal_file_system_directory_ops/test_filesystem_renamefile
// origin: languages/pascal/tests/pascal/test_pascal_file_system_directory_ops.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: TextFile;
begin
  AssignFile(f, 'test_rename_1.tmp'); Rewrite(f); WriteLn(f, 'data'); CloseFile(f);
  RenameFile('test_rename_1.tmp', 'test_rename_2.tmp');
  WriteLn(FileExists('test_rename_1.tmp'));
  WriteLn(FileExists('test_rename_2.tmp'));
  DeleteFile('test_rename_2.tmp');
end.
