// vybe-test: pascal/pascal_typed_file_io/test_typed_file_erase
// origin: languages/pascal/tests/pascal/test_pascal_typed_file_io.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: file of Integer; val: Integer;
begin
  AssignFile(f, 'test_erase.bin');
  Rewrite(f);
  val := 42; Write(f, val);
  CloseFile(f);
  Erase(f);
  WriteLn('ErasedTypedFile');
end.
