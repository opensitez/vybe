program ProceduresFunctions;

function Add(A, B: Integer): Integer;
begin
  Result := A + B;
end;

function Factorial(N: Integer): Integer;
begin
  if N <= 1 then
    Result := 1
  else
    Result := N * Factorial(N - 1);
end;

function IsPrime(N: Integer): Boolean;
var
  I: Integer;
begin
  if N < 2 then
  begin
    Result := False;
    Exit;
  end;
  for I := 2 to N - 1 do
    if N mod I = 0 then
    begin
      Result := False;
      Exit;
    end;
  Result := True;
end;

procedure PrintRange(Lo, Hi: Integer);
var
  I: Integer;
begin
  for I := Lo to Hi do
    Write(I, ' ');
  Writeln;
end;

function Power(Base: Real; Exp: Integer): Real;
var
  I: Integer;
begin
  Result := 1;
  for I := 1 to Exp do
    Result := Result * Base;
end;

begin
  Writeln('Add(3,4) = ', Add(3, 4));
  Writeln('Factorial(5) = ', Factorial(5));
  Writeln('IsPrime(17) = ', IsPrime(17));
  Writeln('IsPrime(18) = ', IsPrime(18));
  PrintRange(1, 10);
  Writeln('2^10 = ', Power(2, 10):0:0);
end.
