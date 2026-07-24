use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 10: Nested Routine Lexical Scoping
// ═══════════════════════════════════════════════════════════

#[test]
fn test_nested_procedure_reads_outer_local() {
    let out = run_pascal(r#"
program Test;
procedure Outer;
var outerVal: Integer;
  procedure Inner;
  begin
    WriteLn(outerVal);
  end;
begin
  outerVal := 42;
  Inner;
end;
begin
  Outer;
end.
"#);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn test_nested_procedure_mutates_outer_local() {
    let out = run_pascal(r#"
program Test;
procedure Counter;
var count: Integer;
  procedure IncCounter;
  begin
    Inc(count);
  end;
begin
  count := 0;
  IncCounter;
  IncCounter;
  IncCounter;
  WriteLn(count);
end;
begin
  Counter;
end.
"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn test_nested_function_returns_computed_outer_val() {
    let out = run_pascal(r#"
program Test;
function Calculate(base: Integer): Integer;
  function Multiplier: Integer;
  begin
    Result := base * 10;
  end;
begin
  Result := Multiplier + 5;
end;
begin
  WriteLn(Calculate(4));
end.
"#);
    assert_eq!(out, vec!["45"]);
}

#[test]
fn test_three_level_nested_scoping() {
    let out = run_pascal(r#"
program Test;
procedure Level1;
var v1: Integer;
  procedure Level2;
  var v2: Integer;
    procedure Level3;
    begin
      WriteLn(v1 + v2);
    end;
  begin
    v2 := 20;
    Level3;
  end;
begin
  v1 := 10;
  Level2;
end;
begin
  Level1;
end.
"#);
    assert_eq!(out, vec!["30"]);
}

#[test]
fn test_nested_procedure_shadows_outer_variable() {
    let out = run_pascal(r#"
program Test;
procedure Outer;
var x: Integer;
  procedure Inner;
  var x: String;
  begin
    x := 'InnerString';
    WriteLn(x);
  end;
begin
  x := 100;
  Inner;
  WriteLn(x);
end;
begin
  Outer;
end.
"#);
    assert_eq!(out, vec!["InnerString", "100"]);
}

#[test]
fn test_nested_procedure_modifies_outer_string() {
    let out = run_pascal(r#"
program Test;
procedure BuildSentence;
var s: String;
  procedure AppendWord(w: String);
  begin
    s := s + w + ' ';
  end;
begin
  s := '';
  AppendWord('Pascal');
  AppendWord('Is');
  AppendWord('Awesome');
  WriteLn(Trim(s));
end;
begin
  BuildSentence;
end.
"#);
    assert_eq!(out, vec!["Pascal Is Awesome"]);
}

#[test]
fn test_nested_procedure_modifies_outer_array() {
    let out = run_pascal(r#"
program Test;
type TArr = array[1..3] of Integer;
procedure FillArray;
var arr: TArr;
  procedure SetElem(idx, val: Integer);
  begin
    arr[idx] := val;
  end;
begin
  SetElem(1, 10);
  SetElem(2, 20);
  SetElem(3, 30);
  WriteLn(arr[1]);
  WriteLn(arr[2]);
  WriteLn(arr[3]);
end;
begin
  FillArray;
end.
"#);
    assert_eq!(out, vec!["10", "20", "30"]);
}

#[test]
fn test_sibling_nested_procedures() {
    let out = run_pascal(r#"
program Test;
procedure Orchestrator;
var data: Integer;
  procedure StepA;
  begin
    data := 100;
  end;
  procedure StepB;
  begin
    data := data + 50;
  end;
begin
  StepA;
  StepB;
  WriteLn(data);
end;
begin
  Orchestrator;
end.
"#);
    assert_eq!(out, vec!["150"]);
}

#[test]
fn test_nested_procedure_called_in_loop() {
    let out = run_pascal(r#"
program Test;
procedure ProcessLoop;
var sum: Integer;
    i: Integer;
  procedure Accumulate(val: Integer);
  begin
    sum := sum + val;
  end;
begin
  sum := 0;
  for i := 1 to 5 do
    Accumulate(i);
  WriteLn(sum);
end;
begin
  ProcessLoop;
end.
"#);
    assert_eq!(out, vec!["15"]);
}

#[test]
fn test_nested_routine_accesses_outer_record() {
    let out = run_pascal(r#"
program Test;
type TUser = record ID: Integer; Name: String; end;
procedure HandleUser;
var user: TUser;
  procedure UpdateName(newName: String);
  begin
    user.Name := newName;
  end;
begin
  user.ID := 1;
  user.Name := 'OldName';
  UpdateName('NewName');
  WriteLn(user.ID);
  WriteLn(user.Name);
end;
begin
  HandleUser;
end.
"#);
    assert_eq!(out, vec!["1", "NewName"]);
}

#[test]
fn test_nested_routine_accesses_outer_enum() {
    let out = run_pascal(r#"
program Test;
type TStatus = (stPending, stActive, stDone);
procedure Workflow;
var status: TStatus;
  procedure Advance;
  begin
    if status = stPending then status := stActive
    else if status = stActive then status := stDone;
  end;
begin
  status := stPending;
  Advance;
  WriteLn(Ord(status));
  Advance;
  WriteLn(Ord(status));
end;
begin
  Workflow;
end.
"#);
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn test_nested_routine_has_own_local_vars() {
    let out = run_pascal(r#"
program Test;
procedure Outer;
var outerVal: Integer;
  procedure Inner;
  var innerVal: Integer;
  begin
    innerVal := 20;
    WriteLn(outerVal + innerVal);
  end;
begin
  outerVal := 10;
  Inner;
end;
begin
  Outer;
end.
"#);
    assert_eq!(out, vec!["30"]);
}

#[test]
fn test_nested_function_in_class_method() {
    let out = run_pascal(r#"
program Test;
type TWorker = class
  public procedure DoWork;
end;
procedure TWorker.DoWork;
var total: Integer;
  function CalculateStep(step: Integer): Integer;
  begin
    Result := step * 2;
  end;
begin
  total := CalculateStep(5) + CalculateStep(10);
  WriteLn(total);
end;
var w: TWorker;
begin
  w := TWorker.Create;
  w.DoWork;
  w.Free;
end.
"#);
    assert_eq!(out, vec!["30"]);
}

#[test]
fn test_nested_recursive_function() {
    let out = run_pascal(r#"
program Test;
function FactorialSum(n: Integer): Integer;
  function Fact(k: Integer): Integer;
  begin
    if k <= 1 then Result := 1
    else Result := k * Fact(k - 1);
  end;
begin
  Result := Fact(n);
end;
begin
  WriteLn(FactorialSum(5));
end.
"#);
    assert_eq!(out, vec!["120"]);
}

#[test]
fn test_nested_procedure_accesses_outer_var_parameter() {
    let out = run_pascal(r#"
program Test;
procedure OuterWithVar(var refVal: Integer);
  procedure InnerModify;
  begin
    refVal := refVal * 3;
  end;
begin
  InnerModify;
end;
var x: Integer;
begin
  x := 10;
  OuterWithVar(x);
  WriteLn(x);
end.
"#);
    assert_eq!(out, vec!["30"]);
}

#[test]
fn test_nested_procedure_accesses_outer_const() {
    let out = run_pascal(r#"
program Test;
procedure TestConstScope;
const Multiplier = 100;
  procedure InnerPrint(val: Integer);
  begin
    WriteLn(val * Multiplier);
  end;
begin
  InnerPrint(5);
end;
begin
  TestConstScope;
end.
"#);
    assert_eq!(out, vec!["500"]);
}

#[test]
fn test_nested_routine_boolean_accumulator() {
    let out = run_pascal(r#"
program Test;
function CheckAllValid: Boolean;
var allOk: Boolean;
  procedure ValidateItem(cond: Boolean);
  begin
    if not cond then allOk := False;
  end;
begin
  allOk := True;
  ValidateItem(True);
  ValidateItem(True);
  ValidateItem(False);
  Result := allOk;
end;
begin
  WriteLn(CheckAllValid);
end.
"#);
    assert_eq!(out, vec!["False"]);
}

#[test]
fn test_nested_procedure_with_parameters_and_outer_vars() {
    let out = run_pascal(r#"
program Test;
procedure Outer;
var factor: Integer;
  procedure InnerAdd(amount: Integer);
  begin
    factor := factor + amount;
  end;
begin
  factor := 10;
  InnerAdd(5);
  InnerAdd(15);
  WriteLn(factor);
end;
begin
  Outer;
end.
"#);
    assert_eq!(out, vec!["30"]);
}

#[test]
fn test_nested_procedure_pointer_deref_outer() {
    let out = run_pascal(r#"
program Test;
procedure PointerScope;
var val: Integer;
    ptr: PInteger;
  procedure ReadPtr;
  begin
    WriteLn(ptr^);
  end;
begin
  val := 777;
  ptr := @val;
  ReadPtr;
end;
begin
  PointerScope;
end.
"#);
    assert_eq!(out, vec!["777"]);
}

#[test]
fn test_nested_function_condition_evaluation() {
    let out = run_pascal(r#"
program Test;
function FilterData(x: Integer): Boolean;
var limit: Integer;
  function IsAboveLimit: Boolean;
  begin
    Result := x > limit;
  end;
begin
  limit := 50;
  Result := IsAboveLimit;
end;
begin
  WriteLn(FilterData(75));
  WriteLn(FilterData(25));
end.
"#);
    assert_eq!(out, vec!["True", "False"]);
}
