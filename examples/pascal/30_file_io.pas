program FileIODemo;

var
  F: TextFile;
  Line: string;
  I: Integer;
begin
  AssignFile(F, '/tmp/test_output.txt');
  Rewrite(F);
  for I := 1 to 5 do
    Writeln(F, 'Line ', I);
  CloseFile(F);

  AssignFile(F, '/tmp/test_output.txt');
  Reset(F);
  Writeln('File contents:');
  while not Eof(F) do
  begin
    Readln(F, Line);
    Writeln(Line);
  end;
  CloseFile(F);
end.
