program GenericsDemo;

type
  TBox<T> = class
  private
    FValue: T;
  public
    constructor Create(AValue: T);
    function GetValue: T;
    procedure SetValue(AValue: T);
    property Value: T read GetValue write SetValue;
  end;

  TPair<TKey, TValue> = class
  private
    FKey: TKey;
    FValue: TValue;
  public
    constructor Create(const AKey: TKey; const AValue: TValue);
    property Key: TKey read FKey;
    property Value: TValue read FValue;
  end;

constructor TBox<T>.Create(AValue: T);
begin
  FValue := AValue;
end;

function TBox<T>.GetValue: T;
begin
  Result := FValue;
end;

procedure TBox<T>.SetValue(AValue: T);
begin
  FValue := AValue;
end;

constructor TPair<TKey, TValue>.Create(const AKey: TKey; const AValue: TValue);
begin
  FKey := AKey;
  FValue := AValue;
end;

function Max<T>(const A, B: T): T;
begin
  if A > B then
    Result := A
  else
    Result := B;
end;

var
  IntBox: TBox<Integer>;
  StrBox: TBox<string>;
  Pair: TPair<string, Integer>;
begin
  IntBox := TBox<Integer>.Create(42);
  Writeln('IntBox = ', IntBox.Value);
  IntBox.Value := 100;
  Writeln('IntBox = ', IntBox.Value);

  StrBox := TBox<string>.Create('Hello');
  Writeln('StrBox = ', StrBox.Value);

  Pair := TPair<string, Integer>.Create('Age', 25);
  Writeln('Pair: ', Pair.Key, ' = ', Pair.Value);

  Writeln('Max(3,7) = ', Max<Integer>(3, 7));
  Writeln('Max(3.5,2.1) = ', Max<Real>(3.5, 2.1):0:1);

  IntBox.Free;
  StrBox.Free;
  Pair.Free;
end.
