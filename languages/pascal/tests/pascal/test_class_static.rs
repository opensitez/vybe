/// Tests for class-level (static) members in Pascal/Delphi:
/// class function, class procedure, class const, class var,
/// and class-wide behavior shared across all instances.
use super::helpers::run_pascal;

// ===================================================================
// CLASS FUNCTION (static method returning value)
// ===================================================================

#[test]
fn class_function_basic() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TMath = class
    class function Square(n: Integer): Integer;
    class function Cube(n: Integer): Integer;
  end;
class function TMath.Square(n: Integer): Integer;
begin
  Result := n * n;
end;
class function TMath.Cube(n: Integer): Integer;
begin
  Result := n * n * n;
end;
begin
  WriteLn(TMath.Square(5));
  WriteLn(TMath.Cube(3));
end."#
        ),
        &["25", "27"]
    );
}

#[test]
fn class_function_min_max() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TUtil = class
    class function ClampValue(v, lo, hi: Integer): Integer;
  end;
class function TUtil.ClampValue(v, lo, hi: Integer): Integer;
begin
  if v < lo then Result := lo
  else if v > hi then Result := hi
  else Result := v;
end;
begin
  WriteLn(TUtil.ClampValue(5, 1, 10));
  WriteLn(TUtil.ClampValue(-5, 1, 10));
  WriteLn(TUtil.ClampValue(15, 1, 10));
end."#
        ),
        &["5", "1", "10"]
    );
}

#[test]
fn class_function_returns_string() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TFormatter = class
    class function Repeat(s: String; n: Integer): String;
  end;
class function TFormatter.Repeat(s: String; n: Integer): String;
var i: Integer;
begin
  Result := '';
  for i := 1 to n do
    Result := Result + s;
end;
begin
  WriteLn(TFormatter.Repeat('ab', 3));
end."#
        ),
        &["ababab"]
    );
}

// ===================================================================
// CLASS PROCEDURE (static void method)
// ===================================================================

#[test]
fn class_procedure_basic() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TLogger = class
    class procedure Log(msg: String);
    class procedure Warn(msg: String);
  end;
class procedure TLogger.Log(msg: String);
begin
  WriteLn('[LOG] ' + msg);
end;
class procedure TLogger.Warn(msg: String);
begin
  WriteLn('[WARN] ' + msg);
end;
begin
  TLogger.Log('started');
  TLogger.Warn('low memory');
  TLogger.Log('finished');
end."#
        ),
        &["[LOG] started", "[WARN] low memory", "[LOG] finished"]
    );
}

#[test]
fn class_procedure_prints_array() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TPrinter = class
    class procedure PrintRange(lo, hi: Integer);
  end;
class procedure TPrinter.PrintRange(lo, hi: Integer);
var i: Integer;
begin
  for i := lo to hi do
    WriteLn(i);
end;
begin
  TPrinter.PrintRange(3, 5);
end."#
        ),
        &["3", "4", "5"]
    );
}

// ===================================================================
// CLASS CONST
// ===================================================================

#[test]
fn class_const_usage() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TCircle = class
  const
    PI = 3.14159;
  public
    Radius: Real;
    function Area: Real;
  end;
function TCircle.Area: Real;
begin
  Result := TCircle.PI * Radius * Radius;
end;
var c: TCircle;
begin
  c := TCircle.Create;
  c.Radius := 5.0;
  WriteLn(c.Area > 78.0);
  WriteLn(c.Area < 79.0);
  c.Free;
end."#
        ),
        &["true", "true"]
    );
}

#[test]
fn class_const_in_class_function() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TLimits = class
  const
    MaxItems = 100;
    MinItems = 1;
  public
    class function IsValidCount(n: Integer): Boolean;
  end;
class function TLimits.IsValidCount(n: Integer): Boolean;
begin
  Result := (n >= TLimits.MinItems) and (n <= TLimits.MaxItems);
end;
begin
  WriteLn(TLimits.IsValidCount(50));
  WriteLn(TLimits.IsValidCount(0));
  WriteLn(TLimits.IsValidCount(101));
end."#
        ),
        &["True", "False", "False"]
    );
}

// ===================================================================
// CLASS WITH FACTORY CLASS FUNCTION
// ===================================================================

#[test]
fn factory_via_class_function() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TPoint = class
  public
    X, Y: Real;
    class function Origin: TPoint;
    class function Create(aX, aY: Real): TPoint;
  end;
class function TPoint.Origin: TPoint;
begin
  Result := TPoint.Create;
  Result.X := 0;
  Result.Y := 0;
end;
class function TPoint.Create(aX, aY: Real): TPoint;
begin
  Result := TPoint.Create;
  Result.X := aX;
  Result.Y := aY;
end;
var p: TPoint;
begin
  p := TPoint.Origin;
  WriteLn(p.X);
  WriteLn(p.Y);
  p.Free;
end."#
        ),
        &["0", "0"]
    );
}

// ===================================================================
// CLASS FUNCTION FOR PARSING / CONVERSION
// ===================================================================

#[test]
fn class_function_parse_int() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TParser = class
    class function TryInt(s: String; var n: Integer): Boolean;
  end;
class function TParser.TryInt(s: String; var n: Integer): Boolean;
begin
  try
    n := StrToInt(s);
    Result := True;
  except
    n := 0;
    Result := False;
  end;
end;
var n: Integer;
begin
  WriteLn(TParser.TryInt('42', n));
  WriteLn(n);
  WriteLn(TParser.TryInt('bad', n));
  WriteLn(n);
end."#
        ),
        &["True", "42", "False", "0"]
    );
}

// ===================================================================
// INHERITED CLASS WITH CLASS FUNCTION
// ===================================================================

#[test]
fn class_function_in_derived() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TAnimal = class
    class function Species: String; virtual;
  end;
  TDog = class(TAnimal)
    class function Species: String; override;
  end;
class function TAnimal.Species: String;
begin
  Result := 'Animal';
end;
class function TDog.Species: String;
begin
  Result := 'Dog';
end;
begin
  WriteLn(TAnimal.Species);
  WriteLn(TDog.Species);
end."#
        ),
        &["Animal", "Dog"]
    );
}

// ===================================================================
// CLASS FUNCTION COMBINED WITH INSTANCE METHODS
// ===================================================================

#[test]
fn class_and_instance_methods_together() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TCounter = class
  private
    FValue: Integer;
  public
    class function Zero: TCounter;
    procedure Increment;
    function Value: Integer;
  end;
class function TCounter.Zero: TCounter;
begin
  Result := TCounter.Create;
  Result.FValue := 0;
end;
procedure TCounter.Increment;
begin
  FValue := FValue + 1;
end;
function TCounter.Value: Integer;
begin
  Result := FValue;
end;
var c: TCounter;
begin
  c := TCounter.Zero;
  c.Increment;
  c.Increment;
  c.Increment;
  WriteLn(c.Value);
  c.Free;
end."#
        ),
        &["3"]
    );
}

// ===================================================================
// MULTIPLE CLASS FUNCTIONS, COMPOSITION
// ===================================================================

#[test]
fn class_functions_compose() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TStringUtil = class
    class function Reversed(s: String): String;
    class function IsPalindrome(s: String): Boolean;
  end;
class function TStringUtil.Reversed(s: String): String;
var i: Integer;
begin
  Result := '';
  for i := Length(s) downto 1 do
    Result := Result + s[i];
end;
class function TStringUtil.IsPalindrome(s: String): Boolean;
begin
  Result := s = TStringUtil.Reversed(s);
end;
begin
  WriteLn(TStringUtil.IsPalindrome('racecar'));
  WriteLn(TStringUtil.IsPalindrome('hello'));
end."#
        ),
        &["True", "False"]
    );
}

// ===================================================================
// CLASS PROCEDURE WITH SIDE EFFECTS VIA VAR
// ===================================================================

#[test]
fn class_procedure_output_param() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TCalc = class
    class procedure Divide(a, b: Integer; var q, r: Integer);
  end;
class procedure TCalc.Divide(a, b: Integer; var q, r: Integer);
begin
  q := a div b;
  r := a mod b;
end;
var q, r: Integer;
begin
  TCalc.Divide(17, 5, q, r);
  WriteLn(q);
  WriteLn(r);
end."#
        ),
        &["3", "2"]
    );
}

// ===================================================================
// CLASS FUNCTION RECURSIVE
// ===================================================================

#[test]
fn class_function_recursive() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TFib = class
    class function Compute(n: Integer): Integer;
  end;
class function TFib.Compute(n: Integer): Integer;
begin
  if n <= 1 then Result := n
  else Result := TFib.Compute(n - 1) + TFib.Compute(n - 2);
end;
begin
  WriteLn(TFib.Compute(0));
  WriteLn(TFib.Compute(1));
  WriteLn(TFib.Compute(7));
end."#
        ),
        &["0", "1", "13"]
    );
}

// ===================================================================
// CLASS CONST ACCESS FROM OUTSIDE CLASS
// ===================================================================

#[test]
fn class_const_access_external() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TConfig = class
  const
    Version = '1.0.0';
    MaxConnections = 10;
  end;
begin
  WriteLn(TConfig.Version);
  WriteLn(TConfig.MaxConnections);
end."#
        ),
        &["1.0.0", "10"]
    );
}

// ===================================================================
// CLASS FUNCTION AND INSTANCE INTERPLAY
// ===================================================================

#[test]
fn singleton_via_class_function() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TApp = class
  private
    FName: String;
  public
    class function CreateNamed(name: String): TApp;
    function GetName: String;
  end;
class function TApp.CreateNamed(name: String): TApp;
begin
  Result := TApp.Create;
  Result.FName := name;
end;
function TApp.GetName: String;
begin
  Result := FName;
end;
var app: TApp;
begin
  app := TApp.CreateNamed('MyApp');
  WriteLn(app.GetName);
  app.Free;
end."#
        ),
        &["MyApp"]
    );
}

// ===================================================================
// CLASS FUNCTION RETURNS BOOLEAN PREDICATE
// ===================================================================

#[test]
fn class_function_predicate() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TValidator = class
    class function IsEmail(s: String): Boolean;
    class function IsPositive(n: Integer): Boolean;
  end;
class function TValidator.IsEmail(s: String): Boolean;
begin
  Result := Pos('@', s) > 0;
end;
class function TValidator.IsPositive(n: Integer): Boolean;
begin
  Result := n > 0;
end;
begin
  WriteLn(TValidator.IsEmail('user@example.com'));
  WriteLn(TValidator.IsEmail('notanemail'));
  WriteLn(TValidator.IsPositive(5));
  WriteLn(TValidator.IsPositive(-1));
end."#
        ),
        &["True", "False", "True", "False"]
    );
}
