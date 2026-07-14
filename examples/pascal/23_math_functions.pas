program MathFunctionsDemo;

var
  X, Y: Real;
  N: Integer;
begin
  X := 16;
  Writeln('sqrt(16) = ', Sqrt(X):0:2);
  Writeln('abs(-5.5) = ', Abs(-5.5):0:2);
  Writeln('round(3.7) = ', Round(3.7));
  Writeln('trunc(3.7) = ', Trunc(3.7));
  Writeln('floor(3.7) = ', Floor(3.7));
  Writeln('ceil(3.2) = ', Ceil(3.2));

  Writeln('sin(pi/2) = ', Sin(Pi / 2):0:4);
  Writeln('cos(0) = ', Cos(0):0:4);
  Writeln('arctan(1) = ', ArcTan(1):0:4);
  Writeln('ln(e) = ', Ln(2.718281828):0:4);
  Writeln('exp(1) = ', Exp(1):0:4);
  Writeln('power(2,10) = ', Power(2, 10):0:0);
  Writeln('sqr(5) = ', Sqr(5):0:0);

  Writeln('min(3,7) = ', Min(3, 7));
  Writeln('max(3,7) = ', Max(3, 7));

  Randomize;
  Writeln('Random = ', Random:0:4);

  N := 5;
  Writeln('succ(5) = ', Succ(N));
  Writeln('pred(5) = ', Pred(N));
  Writeln('inc: '); Inc(N); Writeln(N);
  Writeln('dec: '); Dec(N); Writeln(N);
end.
