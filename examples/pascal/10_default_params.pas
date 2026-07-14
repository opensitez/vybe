program DefaultParamsDemo;

function Greet(Name: string; Greeting: string = 'Hello'; Punct: string = '!'): string;
begin
  Result := Greeting + ', ' + Name + Punct;
end;

function RangeSum(Start: Integer = 1; Finish: Integer = 10): Integer;
var
  I: Integer;
begin
  Result := 0;
  for I := Start to Finish do
    Result := Result + I;
end;

procedure Log(Message: string; Level: Integer = 1; Prefix: string = '[INFO]');
begin
  Writeln(Prefix, ' ', Message, ' (level=', Level, ')');
end;

begin
  Writeln(Greet('Alice'));
  Writeln(Greet('Bob', 'Hi'));
  Writeln(Greet('Carol', 'Welcome', '.'));

  Writeln('RangeSum() = ', RangeSum);
  Writeln('RangeSum(5) = ', RangeSum(5));
  Writeln('RangeSum(5,15) = ', RangeSum(5, 15));

  Log('System started');
  Log('Warning', 2);
  Log('Error', 3, '[ERR]');
end.
