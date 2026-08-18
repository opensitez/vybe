use super::helpers::run_pascal;

#[test]
fn test_scope_global_vs_local() {
    let src = r#"
program T;
var
  x: Integer;

procedure SetLocal;
var
  x: Integer;
begin
  x := 100;
  WriteLn(x);
end;

begin
  x := 5;
  SetLocal;
  WriteLn(x);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["100", "5"]);
}

#[test]
fn test_scope_const_local_hides_global() {
    let src = r#"
program T;
const
  N = 10;

procedure ShowN;
const
  N = 99;
begin
  WriteLn(N);
end;

begin
  WriteLn(N);
  ShowN;
  WriteLn(N);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["10", "99", "10"]);
}

#[test]
fn test_scope_global_var_modified_by_proc() {
    let src = r#"
program T;
var
  total: Integer;

procedure AddToTotal(n: Integer);
begin
  total := total + n;
end;

begin
  total := 0;
  AddToTotal(5);
  AddToTotal(3);
  AddToTotal(7);
  WriteLn(total);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["15"]);
}

#[test]
fn test_scope_type_local_to_proc() {
    let src = r#"
program T;
procedure UsePair;
type
  TPair = record
    A, B: Integer;
  end;
var
  p: TPair;
begin
  p.A := 3;
  p.B := 4;
  WriteLn(p.A + p.B);
end;

begin
  UsePair;
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["7"]);
}

#[test]
fn test_scope_var_initialized_in_global() {
    let src = r#"
program T;
var
  flag: Boolean = true;
  count: Integer = 0;

begin
  WriteLn(flag);
  WriteLn(count);
  count := count + 1;
  WriteLn(count);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["TRUE", "0", "1"]);
}

#[test]
fn test_scope_function_type_local() {
    let src = r#"
program T;
function Compute: Integer;
type
  TData = record
    Values: array[0..2] of Integer;
  end;
var
  d: TData;
  i, sum: Integer;
begin
  d.Values[0] := 10;
  d.Values[1] := 20;
  d.Values[2] := 30;
  sum := 0;
  for i := 0 to 2 do
    sum := sum + d.Values[i];
  Result := sum;
end;

begin
  WriteLn(Compute);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["60"]);
}

#[test]
fn test_scope_nested_same_name() {
    let src = r#"
program T;
var
  msg: string;

procedure Inner;
var
  msg: string;
begin
  msg := 'inner';
  WriteLn(msg);
end;

procedure Outer;
var
  msg: string;
begin
  msg := 'outer';
  WriteLn(msg);
  Inner;
  WriteLn(msg);
end;

begin
  msg := 'global';
  Outer;
  WriteLn(msg);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["outer", "inner", "outer", "global"]);
}

#[test]
fn test_scope_class_private_field() {
    let src = r#"
program T;
type
  TSecret = class
  private
    FSecret: Integer;
  public
    procedure SetSecret(v: Integer);
    function GetSecret: Integer;
  end;

procedure TSecret.SetSecret(v: Integer);
begin
  FSecret := v;
end;

function TSecret.GetSecret: Integer;
begin
  Result := FSecret;
end;

var
  s: TSecret;
begin
  s := TSecret.Create;
  s.SetSecret(42);
  WriteLn(s.GetSecret);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn test_scope_class_protected_accessible_child() {
    let src = r#"
program T;
type
  TBase = class
  protected
    FValue: Integer;
  public
    procedure SetValue(v: Integer);
  end;
  TChild = class(TBase)
    function GetDouble: Integer;
  end;

procedure TBase.SetValue(v: Integer);
begin
  FValue := v;
end;

function TChild.GetDouble: Integer;
begin
  Result := FValue * 2;
end;

var
  c: TChild;
begin
  c := TChild.Create;
  c.SetValue(21);
  WriteLn(c.GetDouble);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn test_scope_for_loop_var_local() {
    let src = r#"
program T;
var
  i: Integer;
begin
  for i := 1 to 3 do
    WriteLn(i);
  WriteLn(i);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["1", "2", "3", "3"]);
}

#[test]
fn test_scope_const_global_in_function() {
    let src = r#"
program T;
const
  PI = 3.14159;
  MAX = 100;

function CircArea(r: Double): Double;
begin
  Result := PI * r * r;
end;

function Clamp(v: Integer): Integer;
begin
  if v > MAX then Result := MAX
  else Result := v;
end;

begin
  WriteLn(Clamp(150));
  WriteLn(Clamp(50));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["100", "50"]);
}

#[test]
fn test_scope_strict_private_hides_field() {
    let src = r#"
program T;
type
  TCounter = class
  strict private
    FCount: Integer;
  public
    procedure Increment;
    function Value: Integer;
  end;

procedure TCounter.Increment;
begin
  FCount := FCount + 1;
end;

function TCounter.Value: Integer;
begin
  Result := FCount;
end;

var
  c: TCounter;
begin
  c := TCounter.Create;
  c.Increment;
  c.Increment;
  c.Increment;
  WriteLn(c.Value);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn test_scope_multiple_procs_share_global() {
    let src = r#"
program T;
var
  log: string;

procedure Append(s: string);
begin
  if log = '' then log := s
  else log := log + ',' + s;
end;

procedure Step1;
begin
  Append('step1');
end;

procedure Step2;
begin
  Append('step2');
end;

procedure Step3;
begin
  Append('step3');
end;

begin
  log := '';
  Step1;
  Step2;
  Step3;
  WriteLn(log);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["step1,step2,step3"]);
}

#[test]
fn test_scope_local_proc_array() {
    let src = r#"
program T;
procedure BuildList;
var
  data: array[1..4] of string;
  i: Integer;
begin
  data[1] := 'a';
  data[2] := 'b';
  data[3] := 'c';
  data[4] := 'd';
  for i := 1 to 4 do
    WriteLn(data[i]);
end;

begin
  BuildList;
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["a", "b", "c", "d"]);
}

#[test]
fn test_scope_var_in_while_block() {
    let src = r#"
program T;
var
  n: Integer;
begin
  n := 1;
  while n <= 5 do begin
    var step: Integer;
    step := n * n;
    WriteLn(step);
    n := n + 1;
  end;
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["1", "4", "9", "16", "25"]);
}

#[test]
fn test_scope_init_block_var() {
    let src = r#"
program T;
var
  a: Integer = 10;
  b: Integer = 20;
  c: Integer = 30;
begin
  WriteLn(a + b + c);
  a := 1; b := 2; c := 3;
  WriteLn(a + b + c);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["60", "6"]);
}

#[test]
fn test_scope_recursive_global_state() {
    let src = r#"
program T;
var
  callCount: Integer;

function Factorial(n: Integer): Integer;
begin
  callCount := callCount + 1;
  if n <= 1 then Result := 1
  else Result := n * Factorial(n - 1);
end;

begin
  callCount := 0;
  WriteLn(Factorial(5));
  WriteLn(callCount);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["120", "5"]);
}

#[test]
fn test_scope_class_constructor_sets_field() {
    let src = r#"
program T;
type
  TPoint = class
  private
    FX, FY: Integer;
  public
    constructor Create(x, y: Integer);
    function ToString: string;
  end;

constructor TPoint.Create(x, y: Integer);
begin
  inherited Create;
  FX := x;
  FY := y;
end;

function TPoint.ToString: string;
begin
  Result := '(' + IntToStr(FX) + ',' + IntToStr(FY) + ')';
end;

var
  p: TPoint;
begin
  p := TPoint.Create(3, 4);
  WriteLn(p.ToString);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["(3,4)"]);
}

#[test]
fn test_scope_shadowing_in_for() {
    let src = r#"
program T;
var
  i: Integer;
begin
  i := 99;
  for i := 1 to 3 do
    Write(IntToStr(i) + ' ');
  WriteLn('');
  WriteLn(i);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["1 2 3 ", "3"]);
}
