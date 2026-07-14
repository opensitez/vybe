program RecursionDemo;

function Factorial(N: Integer): Integer;
begin
  if N <= 1 then
    Result := 1
  else
    Result := N * Factorial(N - 1);
end;

function Fibonacci(N: Integer): Integer;
begin
  if N <= 1 then
    Result := N
  else
    Result := Fibonacci(N - 1) + Fibonacci(N - 2);
end;

function GCD(A, B: Integer): Integer;
begin
  if B = 0 then
    Result := A
  else
    Result := GCD(B, A mod B);
end;

function Power(Base: Real; Exp: Integer): Real;
begin
  if Exp = 0 then
    Result := 1
  else if Exp < 0 then
    Result := 1 / Power(Base, -Exp)
  else
    Result := Base * Power(Base, Exp - 1);
end;

function TowerOfHanoi(N: Integer; FromPeg, ToPeg, AuxPeg: string): Integer;
begin
  if N = 0 then
    Result := 0
  else
  begin
    Result := TowerOfHanoi(N - 1, FromPeg, AuxPeg, ToPeg);
    Writeln('Move disk ', N, ' from ', FromPeg, ' to ', ToPeg);
    Result := Result + 1 + TowerOfHanoi(N - 1, AuxPeg, ToPeg, FromPeg);
  end;
end;

var
  I: Integer;
begin
  Writeln('Factorials:');
  for I := 0 to 10 do
    Writeln(I, '! = ', Factorial(I));

  Writeln('Fibonacci:');
  for I := 0 to 15 do
    Writeln('F(', I, ') = ', Fibonacci(I));

  Writeln('GCD(48, 18) = ', GCD(48, 18));
  Writeln('GCD(56, 98) = ', GCD(56, 98));

  Writeln('2^10 = ', Power(2, 10):0:0);
  Writeln('3^-3 = ', Power(3, -3):0:6);

  Writeln('Tower of Hanoi (3 disks):');
  Writeln('Total moves: ', TowerOfHanoi(3, 'A', 'C', 'B'));
end.
