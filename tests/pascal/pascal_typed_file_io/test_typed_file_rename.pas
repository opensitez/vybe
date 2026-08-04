// vybe-test: pascal/pascal_typed_file_io/test_typed_file_rename
// origin: languages/pascal/tests/pascal/test_pascal_typed_file_io.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: file of Integer; val: Integer;
begin
  AssignFile(f, 'test_orig.bin');
  Rewrite(f);
  val := 55; Write(f, val);
  CloseFile(f);
  Rename(f, 'test_renamed.bin');
  WriteLn('RenamedTypedFile');
end.
