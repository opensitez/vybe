// vybe-test: pascal/pascal_typed_file_io/test_typed_file_in_place_update
// origin: languages/pascal/tests/pascal/test_pascal_typed_file_io.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: file of Integer; val: Integer;
begin
  AssignFile(f, 'test_update.bin');
  Rewrite(f);
  val := 100; Write(f, val);
  val := 200; Write(f, val);
  CloseFile(f);

  Reset(f);
  Seek(f, 1);
  Read(f, val);
  val := val + 50;
  Seek(f, 1);
  Write(f, val);
  CloseFile(f);

  Reset(f);
  Seek(f, 1);
  Read(f, val);
  WriteLn(val);
  CloseFile(f);
end.
