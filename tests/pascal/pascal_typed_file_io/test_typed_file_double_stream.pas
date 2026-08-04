// vybe-test: pascal/pascal_typed_file_io/test_typed_file_double_stream
// origin: languages/pascal/tests/pascal/test_pascal_typed_file_io.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: file of Real; r: Real;
begin
  AssignFile(f, 'test_real.bin');
  Rewrite(f);
  r := 12.34; Write(f, r);
  CloseFile(f);

  Reset(f);
  Read(f, r); WriteLn(r);
  CloseFile(f);
end.
