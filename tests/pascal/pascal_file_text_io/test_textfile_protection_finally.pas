// vybe-test: pascal/pascal_file_text_io/test_textfile_protection_finally
// origin: languages/pascal/tests/pascal/test_pascal_file_text_io.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: TextFile;
begin
  AssignFile(f, 'test_finally.txt');
  Rewrite(f);
  try
    WriteLn(f, 'ProtectedFileContent');
  finally
    CloseFile(f);
    WriteLn('ClosedInFinally');
  end;
end.
