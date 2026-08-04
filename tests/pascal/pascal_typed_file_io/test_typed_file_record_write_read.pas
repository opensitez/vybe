// vybe-test: pascal/pascal_typed_file_io/test_typed_file_record_write_read
// origin: languages/pascal/tests/pascal/test_pascal_typed_file_io.rs
program Test;
{$mode delphi}
uses SysUtils;
type TCustomer = packed record
  ID: Integer;
  Score: Word;
end;
var f: file of TCustomer; c1, c2: TCustomer;
begin
  AssignFile(f, 'test_cust.bin');
  Rewrite(f);
  c1.ID := 1; c1.Score := 95; Write(f, c1);
  c1.ID := 2; c1.Score := 88; Write(f, c1);
  CloseFile(f);

  Reset(f);
  Read(f, c2); WriteLn(c2.ID.ToString + ':' + c2.Score.ToString);
  Read(f, c2); WriteLn(c2.ID.ToString + ':' + c2.Score.ToString);
  CloseFile(f);
end.
