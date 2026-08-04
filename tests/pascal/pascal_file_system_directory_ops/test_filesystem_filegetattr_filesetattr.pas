// vybe-test: pascal/pascal_file_system_directory_ops/test_filesystem_filegetattr_filesetattr
// origin: languages/pascal/tests/pascal/test_pascal_file_system_directory_ops.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: TextFile; fname: String; attr: Integer;
begin
  fname := 'test_attr.tmp';
  AssignFile(f, fname); Rewrite(f); CloseFile(f);

  attr := FileGetAttr(fname);
  WriteLn(attr <> -1);

  DeleteFile(fname);
end.
