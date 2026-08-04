// vybe-test: pascal/pascal_untyped_file_blockread/test_untyped_file_procedure_buffer_parameter
// origin: languages/pascal/tests/pascal/test_pascal_untyped_file_blockread.rs
program Test;
{$mode delphi}
uses SysUtils;
procedure WriteRawData(var f: file; const buffer; size: Integer);
var written: Integer;
begin
  BlockWrite(f, buffer, size, written);
end;
var f: file; val: Integer; readVal: Integer; readBytes: Integer;
begin
  val := 999;
  AssignFile(f, 'test_raw_param.bin');
  Rewrite(f, 1);
  WriteRawData(f, val, SizeOf(Integer));
  CloseFile(f);

  Reset(f, 1);
  BlockRead(f, readVal, SizeOf(Integer), readBytes);
  WriteLn(readVal);
  CloseFile(f);
end.
