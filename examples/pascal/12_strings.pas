program StringsDemo;

var
  S1, S2, S3: string;
  I: Integer;
begin
  S1 := 'Hello, Pascal!';
  Writeln('Original: ', S1);
  Writeln('Length: ', Length(S1));
  Writeln('Uppercase: ', UpCase(S1));
  Writeln('Lowercase: ', LowerCase(S1));
  Writeln('Trim: ', Trim('  spaces  '));

  S2 := Copy(S1, 8, 6);
  Writeln('Copy(8,6): ', S2);

  Writeln('Pos(''Pascal''): ', Pos('Pascal', S1));

  S3 := 'Delphi';
  Writeln('LeftStr(3): ', LeftStr(S3, 3));
  Writeln('RightStr(3): ', RightStr(S3, 3));

  S3 := StringOfChar('*', 10);
  Writeln('Stars: ', S3);

  S3 := 'apple,banana,cherry';
  Writeln('Replace: ', StringReplace(S3, 'banana', 'grape', [rfReplaceAll]));

  S1 := 'A';
  S2 := 'B';
  Writeln('Compare: ', CompareStr(S1, S2));

  for I := 1 to Length(S3) do
    Write(S3[I]);
  Writeln;
end.
