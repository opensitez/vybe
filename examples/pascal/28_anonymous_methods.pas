program AnonymousMethodsDemo;

type
  TIntFunc = function(X: Integer): Integer;
  TIntPredicate = function(X: Integer): Boolean;
  TProc = procedure;

function ApplyTwice(F: TIntFunc; X: Integer): Integer;
begin
  Result := F(F(X));
end;

function MakeMultiplier(Factor: Integer): TIntFunc;
begin
  Result := function(X: Integer): Integer
  begin
    Result := X * Factor;
  end;
end;

function FilterArray(const Arr: array of Integer; Pred: TIntPredicate): array of Integer;
var
  I, Count: Integer;
begin
  Count := 0;
  for I := 0 to High(Arr) do
    if Pred(Arr[I]) then
    begin
      SetLength(Result, Count + 1);
      Result[Count] := Arr[I];
      Count := Count + 1;
    end;
end;

var
  DoubleFn, TripleFn: TIntFunc;
  Numbers, Evens: array of Integer;
  I: Integer;
begin
  DoubleFn := function(X: Integer): Integer
  begin
    Result := X * 2;
  end;

  Writeln('DoubleFn(5) = ', DoubleFn(5));
  Writeln('ApplyTwice(double, 3) = ', ApplyTwice(DoubleFn, 3));

  TripleFn := MakeMultiplier(3);
  Writeln('TripleFn(4) = ', TripleFn(4));

  SetLength(Numbers, 6);
  Numbers[0] := 1;
  Numbers[1] := 2;
  Numbers[2] := 3;
  Numbers[3] := 4;
  Numbers[4] := 5;
  Numbers[5] := 6;

  Evens := FilterArray(Numbers, function(X: Integer): Boolean
  begin
    Result := X mod 2 = 0;
  end);

  Writeln('Evens:');
  for I := 0 to High(Evens) do
    Write(Evens[I], ' ');
  Writeln;
end.
