use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 82: Anonymous Functions, Closures & Lexical Scope
// ═══════════════════════════════════════════════════════════

#[test]
fn test_anonymous_func_basic_inline_invocation() {
    let out = run_pascal(r#"
program Test;
type TFunc = reference to function(a, b: Integer): Integer;
var add: TFunc;
begin
  add := function(a, b: Integer): Integer
  begin
    Result := a + b;
  end;
  WriteLn(add(10, 20));
end.
"#);
    assert_eq!(out, vec!["30"]);
}

#[test]
fn test_anonymous_proc_basic_parameter() {
    let out = run_pascal(r#"
program Test;
type TPrinter = reference to procedure(const s: String);
procedure ExecutePrinter(printer: TPrinter);
begin
  printer('AnonymousProcedureOutput');
end;
begin
  ExecutePrinter(procedure(const s: String)
  begin
    WriteLn(s);
  end);
end.
"#);
    assert_eq!(out, vec!["AnonymousProcedureOutput"]);
}

#[test]
fn test_anonymous_closure_variable_capture() {
    let out = run_pascal(r#"
program Test;
type TIncrementer = reference to function: Integer;
function MakeCounter(initial: Integer): TIncrementer;
var count: Integer;
begin
  count := initial;
  Result := function: Integer
  begin
    Inc(count);
    Result := count;
  end;
end;
var counter: TIncrementer;
begin
  counter := MakeCounter(10);
  WriteLn(counter());
  WriteLn(counter());
end.
"#);
    assert_eq!(out, vec!["11", "12"]);
}

#[test]
fn test_anonymous_method_factory_pattern() {
    let out = run_pascal(r#"
program Test;
type TMultiplier = reference to function(x: Integer): Integer;
function CreateMultiplier(factor: Integer): TMultiplier;
begin
  Result := function(x: Integer): Integer
  begin
    Result := x * factor;
  end;
end;
var mult3, mult5: TMultiplier;
begin
  mult3 := CreateMultiplier(3);
  mult5 := CreateMultiplier(5);
  WriteLn(mult3(10));
  WriteLn(mult5(10));
end.
"#);
    assert_eq!(out, vec!["30", "50"]);
}

#[test]
fn test_anonymous_func_in_tlist() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
type TTask = reference to procedure;
var list: TList<TTask>; task: TTask;
begin
  list := TList<TTask>.Create;
  list.Add(procedure begin WriteLn('Task1'); end);
  list.Add(procedure begin WriteLn('Task2'); end);

  for task in list do task();

  list.Free;
end.
"#);
    assert_eq!(out, vec!["Task1", "Task2"]);
}

#[test]
fn test_anonymous_capturing_string_var() {
    let out = run_pascal(r#"
program Test;
type TStringProc = reference to procedure;
var prefix: String; proc: TStringProc;
begin
  prefix := 'Prefix:';
  proc := procedure
  begin
    WriteLn(prefix + 'Body');
  end;
  prefix := 'UpdatedPrefix:';
  proc();
end.
"#);
    assert_eq!(out, vec!["UpdatedPrefix:Body"]);
}

#[test]
fn test_anonymous_capturing_class_self() {
    let out = run_pascal(r#"
program Test;
type TProc = reference to procedure;
type TMyObj = class
  private FVal: Integer;
  public
    constructor Create(v: Integer);
    function GetRunner: TProc;
end;
constructor TMyObj.Create(v: Integer); begin FVal := v; end;
function TMyObj.GetRunner: TProc;
begin
  Result := procedure
  begin
    WriteLn('CapturedSelfVal:' + FVal.ToString);
  end;
end;

var obj: TMyObj; runner: TProc;
begin
  obj := TMyObj.Create(99);
  runner := obj.GetRunner;
  runner();
  obj.Free;
end.
"#);
    assert_eq!(out, vec!["CapturedSelfVal:99"]);
}

#[test]
fn test_anonymous_nested_closure() {
    let out = run_pascal(r#"
program Test;
type TOuter = reference to function: reference to function: Integer;
var outerFn: TOuter; innerFn: reference to function: Integer;
begin
  outerFn := function: reference to function: Integer
  var x: Integer;
  begin
    x := 100;
    Result := function: Integer
    begin
      Inc(x);
      Result := x;
    end;
  end;

  innerFn := outerFn();
  WriteLn(innerFn());
  WriteLn(innerFn());
end.
"#);
    assert_eq!(out, vec!["101", "102"]);
}

#[test]
fn test_anonymous_function_predicate_filter() {
    let out = run_pascal(r#"
program Test;
type TPredicate = reference to function(x: Integer): Boolean;
procedure FilterAndPrint(const arr: array of Integer; pred: TPredicate);
var i: Integer;
begin
  for i := Low(arr) to High(arr) do
    if pred(arr[i]) then WriteLn(arr[i]);
end;
begin
  FilterAndPrint([1, 2, 3, 4, 5, 6], function(x: Integer): Boolean
  begin
    Result := x mod 2 = 0;
  end);
end.
"#);
    assert_eq!(out, vec!["2", "4", "6"]);
}

#[test]
fn test_anonymous_exception_handling_inside_body() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
type TSafeRunner = reference to procedure;
var runner: TSafeRunner;
begin
  runner := procedure
  begin
    try
      raise Exception.Create('AnonError');
    except
      on E: Exception do WriteLn('AnonCaught:' + E.Message);
    end;
  end;
  runner();
end.
"#);
    assert_eq!(out, vec!["AnonCaught:AnonError"]);
}

#[test]
fn test_anonymous_capturing_record_struct() {
    let out = run_pascal(r#"
program Test;
type TRec = record Code: Integer; end;
type TRecRunner = reference to procedure;
var r: TRec; runner: TRecRunner;
begin
  r.Code := 555;
  runner := procedure
  begin
    WriteLn(r.Code);
  end;
  runner();
end.
"#);
    assert_eq!(out, vec!["555"]);
}

#[test]
fn test_anonymous_chained_execution() {
    let out = run_pascal(r#"
program Test;
type TTransformer = reference to function(x: Integer): Integer;
function Pipe(f1, f2: TTransformer): TTransformer;
begin
  Result := function(x: Integer): Integer
  begin
    Result := f2(f1(x));
  end;
end;
var doubleIt, addTen, combined: TTransformer;
begin
  doubleIt := function(x: Integer): Integer begin Result := x * 2; end;
  addTen := function(x: Integer): Integer begin Result := x + 10; end;

  combined := Pipe(doubleIt, addTen); // (5 * 2) + 10 = 20
  WriteLn(combined(5));
end.
"#);
    assert_eq!(out, vec!["20"]);
}

#[test]
fn test_anonymous_recursion_via_local_var() {
    let out = run_pascal(r#"
program Test;
type TFact = reference to function(n: Integer): Integer;
var fact: TFact;
begin
  fact := function(n: Integer): Integer
  begin
    if n <= 1 then Result := 1
    else Result := n * fact(n - 1);
  end;
  WriteLn(fact(5));
end.
"#);
    assert_eq!(out, vec!["120"]);
}

#[test]
fn test_anonymous_procedure_no_params() {
    let out = run_pascal(r#"
program Test;
type TSimpleProc = reference to procedure;
var p: TSimpleProc;
begin
  p := procedure
  begin
    WriteLn('NoParamAnon');
  end;
  p();
end.
"#);
    assert_eq!(out, vec!["NoParamAnon"]);
}

#[test]
fn test_anonymous_overloaded_outer_routine() {
    let out = run_pascal(r#"
program Test;
type TIntProc = reference to procedure(v: Integer);
type TStrProc = reference to procedure(v: String);

procedure Exec(p: TIntProc); overload;
begin p(42); end;

procedure Exec(p: TStrProc); overload;
begin p('FortyTwo'); end;

begin
  Exec(procedure(v: Integer) begin WriteLn('Int:' + v.ToString); end);
  Exec(procedure(v: String) begin WriteLn('Str:' + v); end);
end.
"#);
    assert_eq!(out, vec!["Int:42", "Str:FortyTwo"]);
}

#[test]
fn test_anonymous_capturing_multiple_variables() {
    let out = run_pascal(r#"
program Test;
type TCalc = reference to function: Integer;
var a, b, c: Integer; calc: TCalc;
begin
  a := 10; b := 20; c := 30;
  calc := function: Integer
  begin
    Result := a + b + c;
  end;
  WriteLn(calc());
end.
"#);
    assert_eq!(out, vec!["60"]);
}

#[test]
fn test_anonymous_nil_check() {
    let out = run_pascal(r#"
program Test;
type TProc = reference to procedure;
var p: TProc;
begin
  p := nil;
  WriteLn(Assigned(p));
end.
"#);
    assert_eq!(out, vec!["False"]);
}

#[test]
fn test_anonymous_function_returning_boolean() {
    let out = run_pascal(r#"
program Test;
type TCheck = reference to function(x: Integer): Boolean;
var isEven: TCheck;
begin
  isEven := function(x: Integer): Boolean begin Result := x mod 2 = 0; end;
  WriteLn(isEven(4));
  WriteLn(isEven(7));
end.
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_anonymous_capturing_loop_value_snapshot() {
    let out = run_pascal(r#"
program Test;
type TProc = reference to procedure;
var procs: array[0..2] of TProc; i: Integer;
begin
  for i := 0 to 2 do
  begin
    var val: Integer := i * 10;
    procs[i] := procedure
    begin
      WriteLn(val);
    end;
  end;

  procs[0]();
  procs[1]();
  procs[2]();
end.
"#);
    assert_eq!(out, vec!["0", "10", "20"]);
}

#[test]
fn test_anonymous_procedure_in_record_method() {
    let out = run_pascal(r#"
program Test;
type TProc = reference to procedure;
type TRunnerRec = record
  procedure Run(p: TProc);
end;
procedure TRunnerRec.Run(p: TProc);
begin
  p();
end;
var r: TRunnerRec;
begin
  r.Run(procedure begin WriteLn('RunnerRecAnonExecuted'); end);
end.
"#);
    assert_eq!(out, vec!["RunnerRecAnonExecuted"]);
}
