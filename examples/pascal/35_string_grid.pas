program StringGridDemo;

var
  Grid: array of array of string;
  Rows, Cols, I, J: Integer;
begin
  Rows := 3;
  Cols := 4;
  SetLength(Grid, Rows, Cols);

  for I := 0 to Rows - 1 do
    for J := 0 to Cols - 1 do
      Grid[I, J] := 'R' + IntToStr(I) + 'C' + IntToStr(J);

  Writeln('String grid:');
  for I := 0 to Rows - 1 do
  begin
    for J := 0 to Cols - 1 do
      Write(Grid[I, J]:8);
    Writeln;
  end;

  Writeln('Grid[1,2] = ', Grid[1, 2]);
end.
