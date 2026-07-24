use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 9: Forward Declarations & External Specifiers
// ═══════════════════════════════════════════════════════════

#[test]
fn test_forward_procedure_simple() {
    let out = run_pascal(r#"
program Test;
procedure SecondProc; forward;
procedure FirstProc;
begin
  WriteLn('First');
  SecondProc;
end;
procedure SecondProc;
begin
  WriteLn('Second');
end;
begin
  FirstProc;
end.
"#);
    assert_eq!(out, vec!["First", "Second"]);
}

#[test]
fn test_forward_function_with_return_type() {
    let out = run_pascal(r#"
program Test;
function ComputeSum(a, b: Integer): Integer; forward;
procedure Exec;
begin
  WriteLn(ComputeSum(10, 20));
end;
function ComputeSum(a, b: Integer): Integer;
begin
  Result := a + b;
end;
begin
  Exec;
end.
"#);
    assert_eq!(out, vec!["30"]);
}

#[test]
fn test_forward_mutual_recursion_even_odd() {
    let out = run_pascal(r#"
program Test;
function IsOdd(n: Integer): Boolean; forward;
function IsEven(n: Integer): Boolean;
begin
  if n = 0 then Result := True
  else Result := IsOdd(n - 1);
end;
function IsOdd(n: Integer): Boolean;
begin
  if n = 0 then Result := False
  else Result := IsEven(n - 1);
end;
begin
  WriteLn(IsEven(4));
  WriteLn(IsOdd(4));
end.
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_forward_var_parameter() {
    let out = run_pascal(r#"
program Test;
procedure ModifyVal(var x: Integer); forward;
procedure RunTest;
var n: Integer;
begin
  n := 5;
  ModifyVal(n);
  WriteLn(n);
end;
procedure ModifyVal(var x: Integer);
begin
  x := x * 10;
end;
begin
  RunTest;
end.
"#);
    assert_eq!(out, vec!["50"]);
}

#[test]
fn test_forward_out_parameter() {
    let out = run_pascal(r#"
program Test;
procedure GenerateData(out s: String); forward;
procedure RunTest;
var text: String;
begin
  GenerateData(text);
  WriteLn(text);
end;
procedure GenerateData(out s: String);
begin
  s := 'GeneratedContent';
end;
begin
  RunTest;
end.
"#);
    assert_eq!(out, vec!["GeneratedContent"]);
}

#[test]
fn test_forward_const_parameter() {
    let out = run_pascal(r#"
program Test;
function FormatMsg(const msg: String): String; forward;
begin
  WriteLn(FormatMsg('Alert'));
end;
function FormatMsg(const msg: String): String;
begin
  Result := '[MSG]: ' + msg;
end;
"#);
    assert_eq!(out, vec!["[MSG]: Alert"]);
}

#[test]
fn test_forward_nested_routine() {
    let out = run_pascal(r#"
program Test;
procedure Outer;
  procedure Inner2; forward;
  procedure Inner1;
  begin
    WriteLn('Inner1');
    Inner2;
  end;
  procedure Inner2;
  begin
    WriteLn('Inner2');
  end;
begin
  Inner1;
end;
begin
  Outer;
end.
"#);
    assert_eq!(out, vec!["Inner1", "Inner2"]);
}

#[test]
fn test_forward_overloaded_routines() {
    let out = run_pascal(r#"
program Test;
procedure Process(i: Integer); overload; forward;
procedure Process(s: String); overload; forward;
procedure Start;
begin
  Process(123);
  Process('ABC');
end;
procedure Process(i: Integer); overload;
begin
  WriteLn('INT:' + i.ToString);
end;
procedure Process(s: String); overload;
begin
  WriteLn('STR:' + s);
end;
begin
  Start;
end.
"#);
    assert_eq!(out, vec!["INT:123", "STR:ABC"]);
}

#[test]
fn test_forward_record_parameter() {
    let out = run_pascal(r#"
program Test;
type TInfo = record Code: Integer; end;
procedure HandleInfo(info: TInfo); forward;
procedure Trigger;
var i: TInfo;
begin
  i.Code := 200;
  HandleInfo(i);
end;
procedure HandleInfo(info: TInfo);
begin
  WriteLn(info.Code);
end;
begin
  Trigger;
end.
"#);
    assert_eq!(out, vec!["200"]);
}

#[test]
fn test_forward_enum_parameter() {
    let out = run_pascal(r#"
program Test;
type TMode = (mFast, mSlow);
procedure SetMode(mode: TMode); forward;
procedure TestMode;
begin
  SetMode(mFast);
end;
procedure SetMode(mode: TMode);
begin
  WriteLn(Ord(mode));
end;
begin
  TestMode;
end.
"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_forward_cdecl_convention_syntax() {
    let out = run_pascal(r#"
program Test;
procedure CDeclFunc(a: Integer); cdecl; forward;
procedure RunFunc;
begin
  CDeclFunc(42);
end;
procedure CDeclFunc(a: Integer); cdecl;
begin
  WriteLn('CDECL:' + a.ToString);
end;
begin
  RunFunc;
end.
"#);
    assert_eq!(out, vec!["CDECL:42"]);
}

#[test]
fn test_forward_stdcall_convention_syntax() {
    let out = run_pascal(r#"
program Test;
procedure StdCallFunc(a: Integer); stdcall; forward;
procedure RunFunc;
begin
  StdCallFunc(99);
end;
procedure StdCallFunc(a: Integer); stdcall;
begin
  WriteLn('STDCALL:' + a.ToString);
end;
begin
  RunFunc;
end.
"#);
    assert_eq!(out, vec!["STDCALL:99"]);
}

#[test]
fn test_forward_register_convention_syntax() {
    let out = run_pascal(r#"
program Test;
procedure RegFunc(a: Integer); register; forward;
procedure RunFunc;
begin
  RegFunc(77);
end;
procedure RegFunc(a: Integer); register;
begin
  WriteLn('REG:' + a.ToString);
end;
begin
  RunFunc;
end.
"#);
    assert_eq!(out, vec!["REG:77"]);
}

#[test]
fn test_forward_pascal_calling_convention() {
    let out = run_pascal(r#"
program Test;
procedure PasFunc(a: Integer); pascal; forward;
procedure RunFunc;
begin
  PasFunc(88);
end;
procedure PasFunc(a: Integer); pascal;
begin
  WriteLn('PAS:' + a.ToString);
end;
begin
  RunFunc;
end.
"#);
    assert_eq!(out, vec!["PAS:88"]);
}

#[test]
fn test_forward_subrange_parameter() {
    let out = run_pascal(r#"
program Test;
type TScore = 1..100;
procedure ShowScore(s: TScore); forward;
procedure RunScore;
var sc: TScore;
begin
  sc := 95;
  ShowScore(sc);
end;
procedure ShowScore(s: TScore);
begin
  WriteLn(s);
end;
begin
  RunScore;
end.
"#);
    assert_eq!(out, vec!["95"]);
}

#[test]
fn test_forward_array_parameter() {
    let out = run_pascal(r#"
program Test;
type TArr = array[1..3] of Integer;
function SumArray(a: TArr): Integer; forward;
procedure RunSum;
var arr: TArr;
begin
  arr[1] := 10; arr[2] := 20; arr[3] := 30;
  WriteLn(SumArray(arr));
end;
function SumArray(a: TArr): Integer;
begin
  Result := a[1] + a[2] + a[3];
end;
begin
  RunSum;
end.
"#);
    assert_eq!(out, vec!["60"]);
}

#[test]
fn test_forward_multiple_parameters() {
    let out = run_pascal(r#"
program Test;
procedure MultiParam(a: Integer; b: String; c: Boolean); forward;
procedure Execute;
begin
  MultiParam(1, 'Two', True);
end;
procedure MultiParam(a: Integer; b: String; c: Boolean);
begin
  WriteLn(a.ToString + '-' + b + '-' + c.ToString);
end;
begin
  Execute;
end.
"#);
    assert_eq!(out, vec!["1-Two-True"]);
}

#[test]
fn test_forward_default_parameter_in_prototype() {
    let out = run_pascal(r#"
program Test;
procedure DefaultProc(x: Integer = 100); forward;
procedure RunDefault;
begin
  DefaultProc;
  DefaultProc(200);
end;
procedure DefaultProc(x: Integer);
begin
  WriteLn(x);
end;
begin
  RunDefault;
end.
"#);
    assert_eq!(out, vec!["100", "200"]);
}

#[test]
fn test_forward_three_step_chain() {
    let out = run_pascal(r#"
program Test;
procedure Step3; forward;
procedure Step2; forward;
procedure Step1;
begin
  WriteLn('Step1');
  Step2;
end;
procedure Step2;
begin
  WriteLn('Step2');
  Step3;
end;
procedure Step3;
begin
  WriteLn('Step3');
end;
begin
  Step1;
end.
"#);
    assert_eq!(out, vec!["Step1", "Step2", "Step3"]);
}

#[test]
fn test_forward_real_return_function() {
    let out = run_pascal(r#"
program Test;
function GetPi: Real; forward;
procedure PrintPi;
begin
  WriteLn(GetPi);
end;
function GetPi: Real;
begin
  Result := 3.14159;
end;
begin
  PrintPi;
end.
"#);
    assert_eq!(out, vec!["3.14159"]);
}
