// vybe-test: pascal/pascal_file_text_io/test_textfile_erase_file
// origin: languages/pascal/tests/pascal/test_pascal_file_text_io.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: TextFile;
begin
  AssignFile(f, 'test_erase.txt');
  Rewrite(f);
  WriteLn(f, 'Temporary');
  CloseFile(f);
  Erase(f);
  WriteLn('Erased');
end.
