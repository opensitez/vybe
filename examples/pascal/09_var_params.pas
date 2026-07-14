program VarParamsDemo;

procedure Swap(var A, B: Integer);
var
  Temp: Integer;
begin
  Temp := A;
  A := B;
  B := Temp;
end;

procedure Increment(var X: Integer; Delta: Integer);
begin
  X := X + Delta;
end;

procedure GetMinMax(const Arr: array of Integer; var MinVal, MaxVal: Integer);
var
  I: Integer;
begin
  MinVal := Arr[0];
  MaxVal := Arr[0];
  for I := 1 to High(Arr) do
  begin
    if Arr[I] < MinVal then
      MinVal := Arr[I];
    if Arr[I] > MaxVal then
      MaxVal := Arr[I];
  end;
end;

var
  X, Y: Integer;
  Numbers: array[0..4] of Integer;
  MinV, MaxV: Integer;
begin
  X := 10;
  Y := 20;
  Swap(X, Y);
  Writeln('After swap: X=', X, ' Y=', Y);

  Increment(X, 5);
  Writeln('After increment: X=', X);

  Numbers[0] := 23;
  Numbers[1] := 5;
  Numbers[2] := 67;
  Numbers[3] := 12;
  Numbers[4] := 89;
  GetMinMax(Numbers, MinV, MaxV);
  Writeln('Min=', MinV, ' Max=', MaxV);
end.
