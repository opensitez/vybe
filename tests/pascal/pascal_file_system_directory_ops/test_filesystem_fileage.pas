// vybe-test: pascal/pascal_file_system_directory_ops/test_filesystem_fileage
// origin: languages/pascal/tests/pascal/test_pascal_file_system_directory_ops.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: TextFile; fname: String; age: Integer;
begin
  fname := 'test_age.tmp';
  AssignFile(f, fname); Rewrite(f); CloseFile(f);

  age := FileAge(fname);
  WriteLn(age <> -1);

  DeleteFile(fname);
end.
