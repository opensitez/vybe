// vybe-test: pascal/pascal_typed_file_io/test_typed_file_empty_check
// origin: languages/pascal/tests/pascal/test_pascal_typed_file_io.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: file of Integer;
begin
  AssignFile(f, 'test_empty.bin');
  Rewrite(f);
  CloseFile(f);

  Reset(f);
  WriteLn(FileSize(f) = 0);
  WriteLn(Eof(f));
  CloseFile(f);
end.
