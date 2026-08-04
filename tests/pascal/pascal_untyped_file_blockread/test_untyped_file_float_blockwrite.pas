// vybe-test: pascal/pascal_untyped_file_blockread/test_untyped_file_float_blockwrite
// origin: languages/pascal/tests/pascal/test_pascal_untyped_file_blockread.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: file; r, readR: Real; written, readBytes: Integer;
begin
  r := 98.76;
  AssignFile(f, 'test_real_raw.bin');
  Rewrite(f, 1);
  BlockWrite(f, r, SizeOf(Real), written);
  CloseFile(f);

  Reset(f, 1);
  BlockRead(f, readR, SizeOf(Real), readBytes);
  WriteLn(readR);
  CloseFile(f);
end.
