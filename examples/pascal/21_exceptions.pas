program ExceptionsDemo;

function SafeDivide(A, B: Real): Real;
begin
  if B = 0 then
    raise Exception.Create('Division by zero');
  Result := A / B;
end;

function ParsePositiveInt(const S: string): Integer;
var
  N: Integer;
begin
  N := StrToInt(S);
  if N < 0 then
    raise Exception.Create('Expected positive number');
  Result := N;
end;

procedure NestedTryDemo;
var
  X: Integer;
begin
  try
    try
      X := StrToInt('not_a_number');
    except
      on E: Exception do
        Writeln('Inner caught: ', E.Message);
    end;
    Writeln('After inner');
    X := SafeDivide(10, 0);
  except
    on E: Exception do
      Writeln('Outer caught: ', E.Message);
  end;
end;

procedure FinallyDemo;
var
  Value: Integer;
begin
  Value := 0;
  try
    Value := 100;
    Writeln('In try, value = ', Value);
  finally
    Value := 0;
    Writeln('In finally, value reset = ', Value);
  end;
end;

var
  ResultVal: Real;
begin
  try
    ResultVal := SafeDivide(100, 5);
    Writeln('100/5 = ', ResultVal:0:2);

    ResultVal := SafeDivide(100, 0);
    Writeln('Should not reach here');
  except
    on E: Exception do
      Writeln('Caught: ', E.Message);
  end;

  try
    Writeln('Parsed: ', ParsePositiveInt('42'));
    Writeln('Parsed: ', ParsePositiveInt('-5'));
  except
    on E: Exception do
      Writeln('Caught: ', E.Message);
  end;

  NestedTryDemo;
  FinallyDemo;
end.
