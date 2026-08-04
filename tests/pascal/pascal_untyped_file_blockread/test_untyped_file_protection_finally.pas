// vybe-test: pascal/pascal_untyped_file_blockread/test_untyped_file_protection_finally
// origin: languages/pascal/tests/pascal/test_pascal_untyped_file_blockread.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: file; b: Byte; written: Integer;
begin
  b := 42;
  AssignFile(f, 'test_prot_untyped.bin');
  Rewrite(f, 1);
  try
    BlockWrite(f, b, 1, written);
  finally
    CloseFile(f);
    WriteLn('UntypedFileClosedInFinally');
  end;
end.
