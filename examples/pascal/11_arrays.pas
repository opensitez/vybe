program ArraysDemo;

var
  Numbers: array[0..9] of Integer;
  Matrix: array[0..2, 0..2] of Integer;
  Names: array[1..3] of string;
  DynamicArr: array of Integer;
  I, J: Integer;
begin
  for I := 0 to 9 do
    Numbers[I] := I * I;
  Writeln('Numbers[5] = ', Numbers[5]);

  for I := 0 to 2 do
    for J := 0 to 2 do
      Matrix[I, J] := I * 3 + J;
  Writeln('Matrix[1,2] = ', Matrix[1, 2]);

  Names[1] := 'Alice';
  Names[2] := 'Bob';
  Names[3] := 'Carol';
  for I := 1 to 3 do
    Writeln('Name ', I, ': ', Names[I]);

  SetLength(DynamicArr, 5);
  for I := 0 to 4 do
    DynamicArr[I] := I * 10;
  for I := 0 to High(DynamicArr) do
    Write(DynamicArr[I], ' ');
  Writeln;

  Append(DynamicArr, 99);
  Writeln('After append, last = ', DynamicArr[High(DynamicArr)]);
end.
