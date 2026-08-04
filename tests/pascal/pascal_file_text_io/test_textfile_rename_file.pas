// vybe-test: pascal/pascal_file_text_io/test_textfile_rename_file
// origin: languages/pascal/tests/pascal/test_pascal_file_text_io.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: TextFile;
begin
  AssignFile(f, 'test_orig.txt');
  Rewrite(f);
  WriteLn(f, 'Data');
  CloseFile(f);
  Rename(f, 'test_renamed.txt');
  WriteLn('Renamed');
end.
