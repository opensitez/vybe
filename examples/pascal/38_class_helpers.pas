program ClassHelpersDemo;

type
  TStringHelper = record helper for string
    function Reverse: string;
    function WordCount: Integer;
    function Contains(const Substr: string): Boolean;
  end;

  TIntegerHelper = record helper for Integer
    function IsEven: Boolean;
    function IsOdd: Boolean;
    function Times(Count: Integer): Integer;
  end;

function TStringHelper.Reverse: string;
var
  I: Integer;
begin
  Result := '';
  for I := Length(Self) downto 1 do
    Result := Result + Self[I];
end;

function TStringHelper.WordCount: Integer;
var
  I: Integer;
  InWord: Boolean;
begin
  Result := 0;
  InWord := False;
  for I := 1 to Length(Self) do
  begin
    if Self[I] <> ' ' then
    begin
      if not InWord then
      begin
        Result := Result + 1;
        InWord := True;
      end;
    end
    else
      InWord := False;
  end;
end;

function TStringHelper.Contains(const Substr: string): Boolean;
begin
  Result := Pos(Substr, Self) > 0;
end;

function TIntegerHelper.IsEven: Boolean;
begin
  Result := Self mod 2 = 0;
end;

function TIntegerHelper.IsOdd: Boolean;
begin
  Result := not IsEven;
end;

function TIntegerHelper.Times(Count: Integer): Integer;
begin
  Result := Self * Count;
end;

var
  S: string;
  N: Integer;
begin
  S := 'Hello World Pascal';
  Writeln('Original: ', S);
  Writeln('Reverse: ', S.Reverse);
  Writeln('WordCount: ', S.WordCount);
  Writeln('Contains ''World'': ', S.Contains('World'));

  N := 42;
  Writeln('IsEven(42): ', N.IsEven);
  Writeln('IsOdd(42): ', N.IsOdd);
  Writeln('42 * 3 = ', N.Times(3));
end.
