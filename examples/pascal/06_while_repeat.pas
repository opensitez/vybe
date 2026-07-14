program WhileRepeatDemo;

var
  N, Factorial: Integer;
  Guess, Target: Integer;
begin
  N := 1;
  Factorial := 1;
  while N <= 7 do
  begin
    Factorial := Factorial * N;
    Writeln(N, '! = ', Factorial);
    N := N + 1;
  end;

  Target := 42;
  Guess := 1;
  repeat
    Guess := Guess + 1;
  until Guess * Guess >= Target;
  Writeln('First square >= ', Target, ' is ', Guess * Guess);

  N := 100;
  while N > 1 do
  begin
    if N mod 2 = 0 then
      N := N div 2
    else
      N := 3 * N + 1;
    Writeln('Collatz: ', N);
  end;
end.
