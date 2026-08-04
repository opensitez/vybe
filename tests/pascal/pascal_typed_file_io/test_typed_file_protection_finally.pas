// vybe-test: pascal/pascal_typed_file_io/test_typed_file_protection_finally
// origin: languages/pascal/tests/pascal/test_pascal_typed_file_io.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: file of Integer; val: Integer;
begin
  AssignFile(f, 'test_prot.bin');
  Rewrite(f);
  try
    val := 777; Write(f, val);
  finally
    CloseFile(f);
    WriteLn('ClosedTypedFileInFinally');
  end;
end.
