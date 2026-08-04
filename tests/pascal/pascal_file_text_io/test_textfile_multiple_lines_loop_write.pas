// vybe-test: pascal/pascal_file_text_io/test_textfile_multiple_lines_loop_write
// origin: languages/pascal/tests/pascal/test_pascal_file_text_io.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: TextFile; i: Integer;
begin
  AssignFile(f, 'test_loop.txt');
  Rewrite(f);
  for i := 1 to 3 do
    WriteLn(f, 'Item ' + i.ToString);
  CloseFile(f);

  Reset(f);
  i := 0;
  while not Eof(f) do
  begin
    Inc(i);
    CloseFile(f); // Break loop early after verification
    Break;
  end;
  WriteLn(i);
end.
