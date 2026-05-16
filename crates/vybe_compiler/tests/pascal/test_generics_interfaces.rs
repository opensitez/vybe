/// Generics, interfaces, operator overloading, class helpers,
/// and other modern Object Pascal / Delphi features.

use super::helpers::run_pascal;

// ===================================================================
// GENERICS — BASIC
// ===================================================================

#[test] fn generic_class_basic() {
    assert_eq!(run_pascal(r#"program T;
type
  TBox<T> = class
  public
    FValue: T;
    constructor Create(v: T);
    function GetValue: T;
  end;

constructor TBox<T>.Create(v: T);
begin FValue := v; end;

function TBox<T>.GetValue: T;
begin Result := FValue; end;

var intBox: TBox<Integer>;
var strBox: TBox<String>;
begin
  intBox := TBox<Integer>.Create(42);
  strBox := TBox<String>.Create('hello');
  WriteLn(intBox.GetValue());
  WriteLn(strBox.GetValue());
end."#), &["42", "hello"]);
}

#[test] fn generic_class_pair() {
    assert_eq!(run_pascal(r#"program T;
type
  TPair<TKey, TValue> = class
  public
    FKey: TKey;
    FValue: TValue;
    constructor Create(k: TKey; v: TValue);
  end;

constructor TPair<TKey, TValue>.Create(k: TKey; v: TValue);
begin FKey := k; FValue := v; end;

var p: TPair<String, Integer>;
begin
  p := TPair<String, Integer>.Create('age', 30);
  WriteLn(p.FKey);
  WriteLn(p.FValue);
end."#), &["age", "30"]);
}

#[test] fn generic_function() {
    assert_eq!(run_pascal(r#"program T;
function Max<T>(a, b: T): T;
begin
  if a > b then Result := a else Result := b;
end;
begin
  WriteLn(Max<Integer>(3, 7));
  WriteLn(Max<String>('alpha', 'beta'));
end."#), &["7", "beta"]);
}

// ===================================================================
// INTERFACES
// ===================================================================

#[test] fn interface_basic_impl() {
    assert_eq!(run_pascal(r#"program T;
type
  IGreeter = interface
    function Greet(name: String): String;
  end;

  TFormalGreeter = class(TInterfacedObject, IGreeter)
  public
    function Greet(name: String): String;
  end;

function TFormalGreeter.Greet(name: String): String;
begin Result := 'Good day, ' + name; end;

var g: IGreeter;
begin
  g := TFormalGreeter.Create;
  WriteLn(g.Greet('Alice'));
end."#), &["Good day, Alice"]);
}

#[test] fn interface_multiple_implementations() {
    assert_eq!(run_pascal(r#"program T;
type
  IShape = interface
    function Area: Real;
  end;
  TSquare = class(TInterfacedObject, IShape)
  public
    FSide: Real;
    constructor Create(s: Real);
    function Area: Real;
  end;
  TCircle = class(TInterfacedObject, IShape)
  public
    FRadius: Real;
    constructor Create(r: Real);
    function Area: Real;
  end;

constructor TSquare.Create(s: Real); begin FSide := s; end;
function TSquare.Area: Real; begin Result := FSide * FSide; end;
constructor TCircle.Create(r: Real); begin FRadius := r; end;
function TCircle.Area: Real; begin Result := 3.14 * FRadius * FRadius; end;

var s: IShape;
begin
  s := TSquare.Create(5);
  WriteLn(s.Area());
  s := TCircle.Create(1);
  WriteLn(s.Area());
end."#), &["25", "3.14"]);
}

// ===================================================================
// OPERATOR OVERLOADING
// ===================================================================

#[test] fn operator_overload_add() {
    assert_eq!(run_pascal(r#"program T;
type
  TVector = record
    X: Real;
    Y: Real;
    class operator Add(a, b: TVector): TVector;
  end;

class operator TVector.Add(a, b: TVector): TVector;
begin
  Result.X := a.X + b.X;
  Result.Y := a.Y + b.Y;
end;

var a, b, c: TVector;
begin
  a.X := 1; a.Y := 2;
  b.X := 3; b.Y := 4;
  c := a + b;
  WriteLn(c.X);
  WriteLn(c.Y);
end."#), &["4", "6"]);
}

#[test] fn operator_overload_equal() {
    assert_eq!(run_pascal(r#"program T;
type
  TPoint = record
    X: Integer;
    Y: Integer;
    class operator Equal(a, b: TPoint): Boolean;
  end;

class operator TPoint.Equal(a, b: TPoint): Boolean;
begin
  Result := (a.X = b.X) and (a.Y = b.Y);
end;

var a, b: TPoint;
begin
  a.X := 1; a.Y := 2;
  b.X := 1; b.Y := 2;
  if a = b then WriteLn('equal') else WriteLn('not equal');
  b.X := 3;
  if a = b then WriteLn('equal') else WriteLn('not equal');
end."#), &["equal", "not equal"]);
}

#[test] fn operator_overload_implicit() {
    assert_eq!(run_pascal(r#"program T;
type
  TWrapper = record
    FValue: Integer;
    class operator Implicit(v: Integer): TWrapper;
    class operator Implicit(w: TWrapper): String;
  end;

class operator TWrapper.Implicit(v: Integer): TWrapper;
begin Result.FValue := v; end;

class operator TWrapper.Implicit(w: TWrapper): String;
begin Result := IntToStr(w.FValue); end;

var w: TWrapper;
var s: String;
begin
  w := 42;
  s := w;
  WriteLn(s);
end."#), &["42"]);
}

// ===================================================================
// CLASS HELPERS
// ===================================================================

#[test] fn class_helper_basic() {
    assert_eq!(run_pascal(r#"program T;
type
  TFoo = class
  public
    FVal: Integer;
    constructor Create(v: Integer);
  end;

  TFooHelper = class helper for TFoo
    function Doubled: Integer;
  end;

constructor TFoo.Create(v: Integer); begin FVal := v; end;
function TFooHelper.Doubled: Integer; begin Result := FVal * 2; end;

var f: TFoo;
begin
  f := TFoo.Create(21);
  WriteLn(f.Doubled());
end."#), &["42"]);
}

// ===================================================================
// RECORD HELPERS
// ===================================================================

#[test] fn record_helper_for_integer() {
    assert_eq!(run_pascal(r#"program T;
type
  TIntHelper = record helper for Integer
    function IsEven: Boolean;
  end;

function TIntHelper.IsEven: Boolean;
begin Result := (Self mod 2) = 0; end;

var n: Integer;
begin
  n := 4;
  WriteLn(n.IsEven());
  n := 7;
  WriteLn(n.IsEven());
end."#), &["true", "false"]);
}

#[test] fn string_helper() {
    assert_eq!(run_pascal(r#"program T;
type
  TStringHelper = record helper for String
    function Reverse: String;
  end;

function TStringHelper.Reverse: String;
var i: Integer;
begin
  Result := '';
  for i := Length(Self) - 1 downto 0 do
    Result := Result + Self[i];
end;

var s: String;
begin
  s := 'hello';
  WriteLn(s.Reverse());
end."#), &["olleh"]);
}

// ===================================================================
// INLINE VARIABLES (Delphi 10.3+)
// ===================================================================

#[test] fn inline_var_in_begin() {
    assert_eq!(run_pascal(r#"program T;
begin
  var x: Integer := 42;
  WriteLn(x);
end."#), &["42"]);
}

#[test] fn inline_var_in_for() {
    assert_eq!(run_pascal(r#"program T;
begin
  var sum: Integer := 0;
  for var i: Integer := 1 to 10 do
    sum := sum + i;
  WriteLn(sum);
end."#), &["55"]);
}

// ===================================================================
// MULTI-INTERFACE
// ===================================================================

#[test] fn class_implements_two_interfaces() {
    assert_eq!(run_pascal(r#"program T;
type
  INameable = interface
    function GetName: String;
  end;
  IDescribable = interface
    function Describe: String;
  end;
  TPet = class(TInterfacedObject, INameable, IDescribable)
  public
    FName: String;
    FKind: String;
    constructor Create(name, kind: String);
    function GetName: String;
    function Describe: String;
  end;

constructor TPet.Create(name, kind: String);
begin FName := name; FKind := kind; end;
function TPet.GetName: String; begin Result := FName; end;
function TPet.Describe: String; begin Result := FName + ' the ' + FKind; end;

var pet: TPet;
var nameable: INameable;
var describable: IDescribable;
begin
  pet := TPet.Create('Rex', 'dog');
  nameable := pet;
  describable := pet;
  WriteLn(nameable.GetName());
  WriteLn(describable.Describe());
end."#), &["Rex", "Rex the dog"]);
}
