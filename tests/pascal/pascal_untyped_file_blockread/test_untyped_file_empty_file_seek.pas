// vybe-test: pascal/pascal_untyped_file_blockread/test_untyped_file_empty_file_seek
// origin: languages/pascal/tests/pascal/test_pascal_untyped_file_blockread.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: file;
begin
  AssignFile(f, 'test_empty_seek.bin');
  Rewrite(f, 1);
  CloseFile(f);

  Reset(f, 1);
  Seek(f, 0);
  WriteLn(FilePos(f));
  CloseFile(f);
end.
