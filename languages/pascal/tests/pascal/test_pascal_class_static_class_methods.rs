use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 15: Class Static Methods & Class Variables
// ═══════════════════════════════════════════════════════════

#[test]
fn test_class_procedure_basic_invocation() {
    let out = run_pascal(r#"
program Test;
type TUtils = class
  public class procedure LogInfo(msg: String);
end;
class procedure TUtils.LogInfo(msg: String);
begin
  WriteLn('[INFO]: ' + msg);
end;
begin
  TUtils.LogInfo('Application Init');
end.
"#);
    assert_eq!(out, vec!["[INFO]: Application Init"]);
}

#[test]
fn test_class_function_return_value() {
    let out = run_pascal(r#"
program Test;
type TMathUtils = class
  public class function Add(a, b: Integer): Integer;
end;
class function TMathUtils.Add(a, b: Integer): Integer;
begin
  Result := a + b;
end;
begin
  WriteLn(TMathUtils.Add(100, 200));
end.
"#);
    assert_eq!(out, vec!["300"]);
}

#[test]
fn test_class_var_shared_state() {
    let out = run_pascal(r#"
program Test;
type TInstanceCounter = class
  public class var Counter: Integer;
  public constructor Create;
end;
constructor TInstanceCounter.Create;
begin
  Inc(Counter);
end;
var i1, i2, i3: TInstanceCounter;
begin
  TInstanceCounter.Counter := 0;
  i1 := TInstanceCounter.Create;
  i2 := TInstanceCounter.Create;
  i3 := TInstanceCounter.Create;
  WriteLn(TInstanceCounter.Counter);
  i1.Free; i2.Free; i3.Free;
end.
"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn test_class_constructor_and_destructor_lifecycle() {
    let out = run_pascal(r#"
program Test;
type TConfig = class
  public class var AppVersion: String;
  class constructor Create;
  class destructor Destroy;
end;
class constructor TConfig.Create;
begin
  AppVersion := 'v1.0.0';
end;
class destructor TConfig.Destroy;
begin
  AppVersion := 'CLEANED';
end;
begin
  WriteLn(TConfig.AppVersion);
end.
"#);
    assert_eq!(out, vec!["v1.0.0"]);
}

#[test]
fn test_class_method_calls_another_class_method() {
    let out = run_pascal(r#"
program Test;
type TStrHelper = class
  public class function Quote(s: String): String;
  public class function DoubleQuote(s: String): String;
end;
class function TStrHelper.Quote(s: String): String;
begin
  Result := "'" + s + "'";
end;
class function TStrHelper.DoubleQuote(s: String): String;
begin
  Result := Quote(Quote(s));
end;
begin
  WriteLn(TStrHelper.DoubleQuote('Text'));
end.
"#);
    assert_eq!(out, vec!["''Text''"]);
}

#[test]
fn test_class_function_factory_pattern() {
    let out = run_pascal(r#"
program Test;
type TProduct = class
  public ID: Integer;
  public class function CreateProduct(AID: Integer): TProduct;
end;
class function TProduct.CreateProduct(AID: Integer): TProduct;
begin
  Result := TProduct.Create;
  Result.ID := AID;
end;
var p: TProduct;
begin
  p := TProduct.CreateProduct(777);
  WriteLn(p.ID);
  p.Free;
end.
"#);
    assert_eq!(out, vec!["777"]);
}

#[test]
fn test_class_procedure_with_default_parameter() {
    let out = run_pascal(r#"
program Test;
type TFormatter = class
  public class procedure PrintHeader(title: String = 'DEFAULT');
end;
class procedure TFormatter.PrintHeader(title: String);
begin
  WriteLn('=== ' + title + ' ===');
end;
begin
  TFormatter.PrintHeader;
  TFormatter.PrintHeader('CUSTOM');
end.
"#);
    assert_eq!(out, vec!["=== DEFAULT ===", "=== CUSTOM ==="]);
}

#[test]
fn test_class_function_overloading() {
    let out = run_pascal(r#"
program Test;
type TParser = class
  public class function Parse(i: Integer): String; overload;
  public class function Parse(s: String): String; overload;
end;
class function TParser.Parse(i: Integer): String; begin Result := 'INT:' + i.ToString; end;
class function TParser.Parse(s: String): String; begin Result := 'STR:' + s; end;
begin
  WriteLn(TParser.Parse(42));
  WriteLn(TParser.Parse('hello'));
end.
"#);
    assert_eq!(out, vec!["INT:42", "STR:hello"]);
}

#[test]
fn test_instance_method_accesses_class_var() {
    let out = run_pascal(r#"
program Test;
type TSession = class
  public class var GlobalTimeout: Integer;
  public function GetTimeout: Integer;
end;
function TSession.GetTimeout: Integer;
begin
  Result := GlobalTimeout;
end;
var s: TSession;
begin
  TSession.GlobalTimeout := 300;
  s := TSession.Create;
  WriteLn(s.GetTimeout);
  s.Free;
end.
"#);
    assert_eq!(out, vec!["300"]);
}

#[test]
fn test_subclass_inherits_class_method() {
    let out = run_pascal(r#"
program Test;
type TBaseUtil = class
  public class procedure BaseLog;
end;
type TSubUtil = class(TBaseUtil) end;
class procedure TBaseUtil.BaseLog;
begin
  WriteLn('BaseLogExecuted');
end;
begin
  TSubUtil.BaseLog;
end.
"#);
    assert_eq!(out, vec!["BaseLogExecuted"]);
}

#[test]
fn test_class_procedure_var_parameter() {
    let out = run_pascal(r#"
program Test;
type TRefUtils = class
  public class procedure MultiplyByTwo(var n: Integer);
end;
class procedure TRefUtils.MultiplyByTwo(var n: Integer);
begin
  n := n * 2;
end;
var val: Integer;
begin
  val := 25;
  TRefUtils.MultiplyByTwo(val);
  WriteLn(val);
end.
"#);
    assert_eq!(out, vec!["50"]);
}

#[test]
fn test_class_method_with_enum_parameter() {
    let out = run_pascal(r#"
program Test;
type TLogLevel = (llInfo, llWarn, llError);
type TLogger = class
  public class procedure Log(lvl: TLogLevel; msg: String);
end;
class procedure TLogger.Log(lvl: TLogLevel; msg: String);
begin
  WriteLn(Ord(lvl).ToString + ':' + msg);
end;
begin
  TLogger.Log(llWarn, 'Disk space low');
end.
"#);
    assert_eq!(out, vec!["1:Disk space low"]);
}

#[test]
fn test_class_function_returning_record() {
    let out = run_pascal(r#"
program Test;
type TPoint = record X, Y: Integer; end;
type TPointFactory = class
  public class function CreatePoint(X, Y: Integer): TPoint;
end;
class function TPointFactory.CreatePoint(X, Y: Integer): TPoint;
begin
  Result.X := X; Result.Y := Y;
end;
var pt: TPoint;
begin
  pt := TPointFactory.CreatePoint(100, 200);
  WriteLn(pt.X);
  WriteLn(pt.Y);
end.
"#);
    assert_eq!(out, vec!["100", "200"]);
}

#[test]
fn test_class_var_private_visibility() {
    let out = run_pascal(r#"
program Test;
type TSecureVault = class
  private class var FMasterKey: Integer;
  public class procedure SetKey(k: Integer);
  public class function GetKey: Integer;
end;
class procedure TSecureVault.SetKey(k: Integer); begin FMasterKey := k; end;
class function TSecureVault.GetKey: Integer; begin Result := FMasterKey; end;
begin
  TSecureVault.SetKey(9876);
  WriteLn(TSecureVault.GetKey);
end.
"#);
    assert_eq!(out, vec!["9876"]);
}

#[test]
fn test_class_function_returning_boolean() {
    let out = run_pascal(r#"
program Test;
type TValidator = class
  public class function IsPositive(n: Integer): Boolean;
end;
class function TValidator.IsPositive(n: Integer): Boolean;
begin
  Result := n > 0;
end;
begin
  WriteLn(TValidator.IsPositive(10));
  WriteLn(TValidator.IsPositive(-5));
end.
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_subclass_overrides_class_method() {
    let out = run_pascal(r#"
program Test;
type TBaseClass = class
  public class procedure Announce; virtual;
end;
type TSubClass = class(TBaseClass)
  public class procedure Announce; override;
end;
class procedure TBaseClass.Announce; begin WriteLn('BaseClass'); end;
class procedure TSubClass.Announce; begin WriteLn('SubClass'); end;
begin
  TBaseClass.Announce;
  TSubClass.Announce;
end.
"#);
    assert_eq!(out, vec!["BaseClass", "SubClass"]);
}

#[test]
fn test_class_function_returning_array() {
    let out = run_pascal(r#"
program Test;
type TIntArr = array[1..3] of Integer;
type TArrayMaker = class
  public class function MakeArray(v1, v2, v3: Integer): TIntArr;
end;
class function TArrayMaker.MakeArray(v1, v2, v3: Integer): TIntArr;
begin
  Result[1] := v1; Result[2] := v2; Result[3] := v3;
end;
var arr: TIntArr;
begin
  arr := TArrayMaker.MakeArray(10, 20, 30);
  WriteLn(arr[2]);
end.
"#);
    assert_eq!(out, vec!["20"]);
}

#[test]
fn test_class_method_called_via_instance_variable() {
    let out = run_pascal(r#"
program Test;
type THelper = class
  public class function Version: String;
end;
class function THelper.Version: String; begin Result := '1.2.3'; end;
var h: THelper;
begin
  h := THelper.Create;
  WriteLn(h.Version);
  h.Free;
end.
"#);
    assert_eq!(out, vec!["1.2.3"]);
}

#[test]
fn test_multiple_class_vars_in_single_declaration() {
    let out = run_pascal(r#"
program Test;
type TStats = class
  public class var Reads, Writes: Integer;
  public class procedure Reset;
end;
class procedure TStats.Reset;
begin
  Reads := 0; Writes := 0;
end;
begin
  TStats.Reads := 10; TStats.Writes := 5;
  TStats.Reset;
  WriteLn(TStats.Reads);
  WriteLn(TStats.Writes);
end.
"#);
    assert_eq!(out, vec!["0", "0"]);
}

#[test]
fn test_class_function_real_precision_computation() {
    let out = run_pascal(r#"
program Test;
type TGeometry = class
  public class function CircleArea(radius: Real): Real;
end;
class function TGeometry.CircleArea(radius: Real): Real;
begin
  Result := 3.14159 * radius * radius;
end;
begin
  WriteLn(TGeometry.CircleArea(2.0));
end.
"#);
    assert_eq!(out, vec!["12.56636"]);
}
