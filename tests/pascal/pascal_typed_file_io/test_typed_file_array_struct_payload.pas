// vybe-test: pascal/pascal_typed_file_io/test_typed_file_array_struct_payload
// origin: languages/pascal/tests/pascal/test_pascal_typed_file_io.rs
program Test;
{$mode delphi}
uses SysUtils;
type TArrayPayload = packed record
  Data: array[0..2] of Integer;
end;
var f: file of TArrayPayload; p1, p2: TArrayPayload;
begin
  AssignFile(f, 'test_arr_payload.bin');
  Rewrite(f);
  p1.Data[0] := 1; p1.Data[1] := 2; p1.Data[2] := 3;
  Write(f, p1);
  CloseFile(f);

  Reset(f);
  Read(f, p2);
  WriteLn(p2.Data[0] + p2.Data[1] + p2.Data[2]);
  CloseFile(f);
end.
