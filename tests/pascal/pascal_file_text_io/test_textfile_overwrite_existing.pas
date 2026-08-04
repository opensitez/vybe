// vybe-test: pascal/pascal_file_text_io/test_textfile_overwrite_existing
// origin: languages/pascal/tests/pascal/test_pascal_file_text_io.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: TextFile; line: String;
begin
  AssignFile(f, 'test_overwrite.txt');
  Rewrite(f); WriteLn(f, 'Old'); CloseFile(f);
  Rewrite(f); WriteLn(f, 'New'); CloseFile(f);

  Reset(f);
  ReadLn(f, line);
  WriteLn(line);
  CloseFile(f);
end.
