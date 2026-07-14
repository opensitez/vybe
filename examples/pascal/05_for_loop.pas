program ForLoopDemo;

var
  I, Sum: Integer;
begin
  Sum := 0;
  for I := 1 to 10 do
    Sum := Sum + I;
  Writeln('Sum 1..10 = ', Sum);

  Sum := 0;
  for I := 10 downto 1 do
    Sum := Sum + I;
  Writeln('Sum 10..1 = ', Sum);

  for I := 0 to 5 do
  begin
    Writeln('Square of ', I, ' = ', I * I);
  end;

  for I := 1 to 20 do
    if I mod 2 = 0 then
      Writeln('Even: ', I);
end.
