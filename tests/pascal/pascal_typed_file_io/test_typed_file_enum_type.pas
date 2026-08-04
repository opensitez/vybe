// vybe-test: pascal/pascal_typed_file_io/test_typed_file_enum_type
// origin: languages/pascal/tests/pascal/test_pascal_typed_file_io.rs
program Test;
{$mode delphi}
uses SysUtils;
type TStatus = (stInit, stRunning, stDone);
var f: file of TStatus; s: TStatus;
begin
  AssignFile(f, 'test_enum.bin');
  Rewrite(f);
  s := stRunning; Write(f, s);
  CloseFile(f);

  Reset(f);
  Read(f, s); WriteLn(Ord(s));
  CloseFile(f);
end.
