// vybe-test: pascal/pascal_typed_file_io/test_typed_file_boolean_type
// origin: languages/pascal/tests/pascal/test_pascal_typed_file_io.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: file of Boolean; b: Boolean;
begin
  AssignFile(f, 'test_bool.bin');
  Rewrite(f);
  b := True; Write(f, b);
  b := False; Write(f, b);
  CloseFile(f);

  Reset(f);
  Read(f, b); WriteLn(b);
  Read(f, b); WriteLn(b);
  CloseFile(f);
end.
