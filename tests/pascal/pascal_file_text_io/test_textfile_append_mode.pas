// vybe-test: pascal/pascal_file_text_io/test_textfile_append_mode
// origin: languages/pascal/tests/pascal/test_pascal_file_text_io.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: TextFile; line: String;
begin
  AssignFile(f, 'test_text_1.txt');
  Append(f);
  WriteLn(f, 'Line 3');
  CloseFile(f);

  Reset(f);
  while not Eof(f) do
  begin
    ReadLn(f, line);
    WriteLn(line);
  end;
  CloseFile(f);
end.
