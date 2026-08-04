// vybe-test: pascal/pascal_typed_file_io/test_typed_file_procedure_parameter
// origin: languages/pascal/tests/pascal/test_pascal_typed_file_io.rs
program Test;
{$mode delphi}
uses SysUtils;
procedure WriteInt(var f: file of Integer; v: Integer);
begin
  Write(f, v);
end;
var f: file of Integer; val: Integer;
begin
  AssignFile(f, 'test_proc_bin.bin');
  Rewrite(f);
  WriteInt(f, 888);
  CloseFile(f);

  Reset(f);
  Read(f, val);
  WriteLn(val);
  CloseFile(f);
end.
