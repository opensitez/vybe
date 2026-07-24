use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 8: Routine & Method Overloading (overload)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_overload_by_parameter_count() {
    let out = run_pascal(r#"
program Test;
procedure Display(a: Integer); overload;
begin
  WriteLn('One: ' + a.ToString);
end;
procedure Display(a, b: Integer); overload;
begin
  WriteLn('Two: ' + a.ToString + ', ' + b.ToString);
end;
begin
  Display(10);
  Display(10, 20);
end.
"#);
    assert_eq!(out, vec!["One: 10", "Two: 10, 20"]);
}

#[test]
fn test_overload_by_parameter_type_int_and_string() {
    let out = run_pascal(r#"
program Test;
procedure PrintVal(v: Integer); overload;
begin
  WriteLn('INT=' + v.ToString);
end;
procedure PrintVal(v: String); overload;
begin
  WriteLn('STR=' + v);
end;
begin
  PrintVal(100);
  PrintVal('Pascal');
end.
"#);
    assert_eq!(out, vec!["INT=100", "STR=Pascal"]);
}

#[test]
fn test_overload_functions_with_different_return_types() {
    let out = run_pascal(r#"
program Test;
function Convert(v: Integer): String; overload;
begin
  Result := 'Int:' + v.ToString;
end;
function Convert(v: Boolean): String; overload;
begin
  if v then Result := 'Bool:True' else Result := 'Bool:False';
end;
begin
  WriteLn(Convert(42));
  WriteLn(Convert(True));
end.
"#);
    assert_eq!(out, vec!["Int:42", "Bool:True"]);
}

#[test]
fn test_overload_class_methods() {
    let out = run_pascal(r#"
program Test;
type TLogger = class
  public procedure Log(msg: String); overload;
  public procedure Log(code: Integer; msg: String); overload;
end;
procedure TLogger.Log(msg: String);
begin
  WriteLn('LOG: ' + msg);
end;
procedure TLogger.Log(code: Integer; msg: String);
begin
  WriteLn('ERR[' + code.ToString + ']: ' + msg);
end;
var logger: TLogger;
begin
  logger := TLogger.Create;
  logger.Log('Info message');
  logger.Log(404, 'Not Found');
  logger.Free;
end.
"#);
    assert_eq!(out, vec!["LOG: Info message", "ERR[404]: Not Found"]);
}

#[test]
fn test_overload_constructors() {
    let out = run_pascal(r#"
program Test;
type TPerson = class
  public Name: String; Age: Integer;
  constructor Create(AName: String); overload;
  constructor Create(AName: String; AAge: Integer); overload;
end;
constructor TPerson.Create(AName: String);
begin
  Name := AName; Age := 0;
end;
constructor TPerson.Create(AName: String; AAge: Integer);
begin
  Name := AName; Age := AAge;
end;
var p1, p2: TPerson;
begin
  p1 := TPerson.Create('Alice');
  p2 := TPerson.Create('Bob', 30);
  WriteLn(p1.Name + ':' + p1.Age.ToString);
  WriteLn(p2.Name + ':' + p2.Age.ToString);
  p1.Free; p2.Free;
end.
"#);
    assert_eq!(out, vec!["Alice:0", "Bob:30"]);
}

#[test]
fn test_overload_class_static_methods() {
    let out = run_pascal(r#"
program Test;
type TMath = class
  public class function Square(x: Integer): Integer; overload;
  public class function Square(x: Real): Real; overload;
end;
class function TMath.Square(x: Integer): Integer;
begin
  Result := x * x;
end;
class function TMath.Square(x: Real): Real;
begin
  Result := x * x;
end;
begin
  WriteLn(TMath.Square(5));
  WriteLn(TMath.Square(2.5));
end.
"#);
    assert_eq!(out, vec!["25", "6.25"]);
}

#[test]
fn test_overload_boolean_vs_integer() {
    let out = run_pascal(r#"
program Test;
procedure Process(b: Boolean); overload;
begin
  WriteLn('BOOL');
end;
procedure Process(i: Integer); overload;
begin
  WriteLn('INT');
end;
begin
  Process(True);
  Process(1);
end.
"#);
    assert_eq!(out, vec!["BOOL", "INT"]);
}

#[test]
fn test_overload_record_vs_pointer() {
    let out = run_pascal(r#"
program Test;
type TRec = record Val: Integer; end;
type PRec = ^TRec;
procedure Inspect(r: TRec); overload;
begin
  WriteLn('VALUE:' + r.Val.ToString);
end;
procedure Inspect(p: PRec); overload;
begin
  WriteLn('POINTER:' + p^.Val.ToString);
end;
var rec: TRec;
begin
  rec.Val := 99;
  Inspect(rec);
  Inspect(@rec);
end.
"#);
    assert_eq!(out, vec!["VALUE:99", "POINTER:99"]);
}

#[test]
fn test_overload_var_vs_value_parameters() {
    let out = run_pascal(r#"
program Test;
procedure UpdateVal(v: Integer); overload;
begin
  WriteLn('VAL:' + v.ToString);
end;
procedure UpdateVal(var v: Integer); overload;
begin
  v := v * 2;
  WriteLn('VAR:' + v.ToString);
end;
var x: Integer;
begin
  x := 10;
  UpdateVal(x);
end.
"#);
    assert_eq!(out, vec!["VAR:20"]);
}

#[test]
fn test_overload_enum_parameters() {
    let out = run_pascal(r#"
program Test;
type TColor = (Red, Green, Blue);
type TSize = (Small, Medium, Large);
procedure Render(c: TColor); overload;
begin
  WriteLn('COLOR:' + Ord(c).ToString);
end;
procedure Render(s: TSize); overload;
begin
  WriteLn('SIZE:' + Ord(s).ToString);
end;
begin
  Render(Green);
  Render(Large);
end.
"#);
    assert_eq!(out, vec!["COLOR:1", "SIZE:2"]);
}

#[test]
fn test_overload_array_types() {
    let out = run_pascal(r#"
program Test;
type TIntArr = array[1..3] of Integer;
type TStrArr = array[1..3] of String;
procedure Dump(a: TIntArr); overload;
begin
  WriteLn('INT_ARR:' + a[1].ToString);
end;
procedure Dump(a: TStrArr); overload;
begin
  WriteLn('STR_ARR:' + a[1]);
end;
var ia: TIntArr; sa: TStrArr;
begin
  ia[1] := 10; sa[1] := 'first';
  Dump(ia);
  Dump(sa);
end.
"#);
    assert_eq!(out, vec!["INT_ARR:10", "STR_ARR:first"]);
}

#[test]
fn test_overload_inherited_class_methods() {
    let out = run_pascal(r#"
program Test;
type TBase = class
  public procedure Show(i: Integer); overload; virtual;
end;
type TDerived = class(TBase)
  public procedure Show(s: String); overload;
end;
procedure TBase.Show(i: Integer);
begin
  WriteLn('BASE_INT:' + i.ToString);
end;
procedure TDerived.Show(s: String);
begin
  WriteLn('DERIVED_STR:' + s);
end;
var d: TDerived;
begin
  d := TDerived.Create;
  d.Show('hello');
  d.Show(42);
  d.Free;
end.
"#);
    assert_eq!(out, vec!["DERIVED_STR:hello", "BASE_INT:42"]);
}

#[test]
fn test_overload_with_default_parameters() {
    let out = run_pascal(r#"
program Test;
procedure Exec(a: Integer; b: Integer = 0); overload;
begin
  WriteLn('EXEC_INT:' + (a + b).ToString);
end;
procedure Exec(s: String; count: Integer = 1); overload;
begin
  WriteLn('EXEC_STR:' + s);
end;
begin
  Exec(10);
  Exec('test');
end.
"#);
    assert_eq!(out, vec!["EXEC_INT:10", "EXEC_STR:test"]);
}

#[test]
fn test_overload_three_types() {
    let out = run_pascal(r#"
program Test;
procedure Handle(i: Integer); overload; begin WriteLn('INT'); end;
procedure Handle(r: Real); overload; begin WriteLn('REAL'); end;
procedure Handle(s: String); overload; begin WriteLn('STR'); end;
begin
  Handle(5);
  Handle(5.5);
  Handle('five');
end.
"#);
    assert_eq!(out, vec!["INT", "REAL", "STR"]);
}

#[test]
fn test_overload_subrange_parameters() {
    let out = run_pascal(r#"
program Test;
type TSub1 = 1..10;
type TSub2 = 100..200;
procedure TestSub(s: TSub1); overload;
begin
  WriteLn('SUB1:' + s.ToString);
end;
procedure TestSub(s: TSub2); overload;
begin
  WriteLn('SUB2:' + s.ToString);
end;
var v1: TSub1; v2: TSub2;
begin
  v1 := 5; v2 := 150;
  TestSub(v1);
  TestSub(v2);
end.
"#);
    assert_eq!(out, vec!["SUB1:5", "SUB2:150"]);
}

#[test]
fn test_overload_in_nested_routines() {
    let out = run_pascal(r#"
program Test;
procedure Parent;
  procedure Action(n: Integer); overload;
  begin
    WriteLn('NESTED_INT:' + n.ToString);
  end;
  procedure Action(s: String); overload;
  begin
    WriteLn('NESTED_STR:' + s);
  end;
begin
  Action(99);
  Action('nested');
end;
begin
  Parent;
end.
"#);
    assert_eq!(out, vec!["NESTED_INT:99", "NESTED_STR:nested"]);
}

#[test]
fn test_overload_char_vs_string() {
    let out = run_pascal(r#"
program Test;
procedure ProcessChar(c: Char); overload;
begin
  WriteLn('CHAR:' + c);
end;
procedure ProcessChar(s: String); overload;
begin
  WriteLn('STRING:' + s);
end;
begin
  ProcessChar('A');
  ProcessChar('ABC');
end.
"#);
    assert_eq!(out, vec!["CHAR:A", "STRING:ABC"]);
}

#[test]
fn test_overload_const_ref_vs_value() {
    let out = run_pascal(r#"
program Test;
type TData = record ID: Integer; end;
procedure LoadData(id: Integer); overload;
begin
  WriteLn('BY_ID:' + id.ToString);
end;
procedure LoadData(const d: TData); overload;
begin
  WriteLn('BY_REC:' + d.ID.ToString);
end;
var rec: TData;
begin
  rec.ID := 55;
  LoadData(55);
  LoadData(rec);
end.
"#);
    assert_eq!(out, vec!["BY_ID:55", "BY_REC:55"]);
}

#[test]
fn test_overload_four_arguments() {
    let out = run_pascal(r#"
program Test;
procedure Config(a, b: Integer); overload;
begin
  WriteLn('2_ARGS:' + (a + b).ToString);
end;
procedure Config(a, b, c, d: Integer); overload;
begin
  WriteLn('4_ARGS:' + (a + b + c + d).ToString);
end;
begin
  Config(1, 2);
  Config(1, 2, 3, 4);
end.
"#);
    assert_eq!(out, vec!["2_ARGS:3", "4_ARGS:10"]);
}

#[test]
fn test_overload_variant_vs_typed() {
    let out = run_pascal(r#"
program Test;
procedure Evaluate(i: Integer); overload;
begin
  WriteLn('TYPED_INT');
end;
procedure Evaluate(v: Variant); overload;
begin
  WriteLn('VARIANT');
end;
var v: Variant;
begin
  Evaluate(42);
  v := 'text';
  Evaluate(v);
end.
"#);
    assert_eq!(out, vec!["TYPED_INT", "VARIANT"]);
}
