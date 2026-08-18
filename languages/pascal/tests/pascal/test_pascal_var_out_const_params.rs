use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 6: Parameter Modifiers (var, out, const, constref)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_var_param_integer_swap() {
    let out = run_pascal(
        r#"
program Test;
procedure SwapInt(var a, b: Integer);
var temp: Integer;
begin
  temp := a;
  a := b;
  b := temp;
end;
var x, y: Integer;
begin
  x := 10;
  y := 20;
  SwapInt(x, y);
  WriteLn(x);
  WriteLn(y);
end.
"#,
    );
    assert_eq!(out, vec!["20", "10"]);
}

#[test]
fn test_out_param_value_assignment() {
    let out = run_pascal(
        r#"
program Test;
procedure GetValues(out a: Integer; out b: String);
begin
  a := 100;
  b := 'Result';
end;
var x: Integer;
    s: String;
begin
  GetValues(x, s);
  WriteLn(x);
  WriteLn(s);
end.
"#,
    );
    assert_eq!(out, vec!["100", "Result"]);
}

#[test]
fn test_const_param_read_only_access() {
    let out = run_pascal(
        r#"
program Test;
function SumConst(const a, b: Integer): Integer;
begin
  Result := a + b;
end;
begin
  WriteLn(SumConst(15, 25));
end.
"#,
    );
    assert_eq!(out, vec!["40"]);
}

#[test]
fn test_var_param_string_mutation() {
    let out = run_pascal(
        r#"
program Test;
procedure AppendTag(var s: String);
begin
  s := s + '_TAGGED';
end;
var text: String;
begin
  text := 'ITEM';
  AppendTag(text);
  WriteLn(text);
end.
"#,
    );
    assert_eq!(out, vec!["ITEM_TAGGED"]);
}

#[test]
fn test_var_param_record_mutation() {
    let out = run_pascal(
        r#"
program Test;
type TPoint = record
  X, Y: Integer;
end;
procedure MovePoint(var p: TPoint; dx, dy: Integer);
begin
  p.X := p.X + dx;
  p.Y := p.Y + dy;
end;
var pt: TPoint;
begin
  pt.X := 5;
  pt.Y := 10;
  MovePoint(pt, 3, 4);
  WriteLn(pt.X);
  WriteLn(pt.Y);
end.
"#,
    );
    assert_eq!(out, vec!["8", "14"]);
}

#[test]
fn test_out_param_record_initialization() {
    let out = run_pascal(
        r#"
program Test;
type TRect = record
  W, H: Integer;
end;
procedure InitRect(out r: TRect; width, height: Integer);
begin
  r.W := width;
  r.H := height;
end;
var myRect: TRect;
begin
  InitRect(myRect, 800, 600);
  WriteLn(myRect.W);
  WriteLn(myRect.H);
end.
"#,
    );
    assert_eq!(out, vec!["800", "600"]);
}

#[test]
fn test_var_param_array_mutation() {
    let out = run_pascal(
        r#"
program Test;
type TIntArray = array[1..3] of Integer;
procedure DoubleElements(var arr: TIntArray);
var i: Integer;
begin
  for i := 1 to 3 do
    arr[i] := arr[i] * 2;
end;
var nums: TIntArray;
begin
  nums[1] := 10; nums[2] := 20; nums[3] := 30;
  DoubleElements(nums);
  WriteLn(nums[1]);
  WriteLn(nums[2]);
  WriteLn(nums[3]);
end.
"#,
    );
    assert_eq!(out, vec!["20", "40", "60"]);
}

#[test]
fn test_constref_parameter_passing() {
    let out = run_pascal(
        r#"
program Test;
type TLargeData = record
  ID: Integer;
  Title: String;
end;
function ReadTitle(constref data: TLargeData): String;
begin
  Result := data.Title;
end;
var d: TLargeData;
begin
  d.ID := 1;
  d.Title := 'Document';
  WriteLn(ReadTitle(d));
end.
"#,
    );
    assert_eq!(out, vec!["Document"]);
}

#[test]
fn test_var_param_passthrough_multiple_levels() {
    let out = run_pascal(
        r#"
program Test;
procedure InnerInc(var val: Integer);
begin
  Inc(val);
end;
procedure OuterInc(var val: Integer);
begin
  InnerInc(val);
  Inc(val);
end;
var n: Integer;
begin
  n := 5;
  OuterInc(n);
  WriteLn(n);
end.
"#,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn test_multiple_var_out_const_in_one_routine() {
    let out = run_pascal(
        r#"
program Test;
procedure ProcessData(const input: String; var count: Integer; out status: String);
begin
  count := count + Length(input);
  status := 'PROCESSED';
end;
var c: Integer;
    st: String;
begin
  c := 10;
  ProcessData('Hello', c, st);
  WriteLn(c);
  WriteLn(st);
end.
"#,
    );
    assert_eq!(out, vec!["15", "PROCESSED"]);
}

#[test]
fn test_var_param_enum_mutation() {
    let out = run_pascal(
        r#"
program Test;
type TState = (StateIdle, StateBusy, StateDone);
procedure AdvanceState(var s: TState);
begin
  if s = StateIdle then s := StateBusy
  else if s = StateBusy then s := StateDone;
end;
var current: TState;
begin
  current := StateIdle;
  AdvanceState(current);
  WriteLn(Ord(current));
  AdvanceState(current);
  WriteLn(Ord(current));
end.
"#,
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn test_var_param_in_recursive_function() {
    let out = run_pascal(
        r#"
program Test;
procedure CountDown(n: Integer; var totalSteps: Integer);
begin
  Inc(totalSteps);
  if n > 1 then
    CountDown(n - 1, totalSteps);
end;
var steps: Integer;
begin
  steps := 0;
  CountDown(5, steps);
  WriteLn(steps);
end.
"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn test_var_param_subrange_type() {
    let out = run_pascal(
        r#"
program Test;
type TScore = 0..100;
procedure AddBonus(var score: TScore; bonus: Integer);
begin
  score := score + bonus;
end;
var myScore: TScore;
begin
  myScore := 80;
  AddBonus(myScore, 15);
  WriteLn(myScore);
end.
"#,
    );
    assert_eq!(out, vec!["95"]);
}

#[test]
fn test_out_param_multiple_return_emulation() {
    let out = run_pascal(
        r#"
program Test;
procedure DivideAndRemainder(dividend, divisor: Integer; out quotient, remainder: Integer);
begin
  quotient := dividend div divisor;
  remainder := dividend mod divisor;
end;
var q, r: Integer;
begin
  DivideAndRemainder(29, 6, q, r);
  WriteLn(q);
  WriteLn(r);
end.
"#,
    );
    assert_eq!(out, vec!["4", "5"]);
}

#[test]
fn test_var_param_pointer_mutation() {
    let out = run_pascal(
        r#"
program Test;
procedure ReassignPointer(var p: PInteger; var target: Integer);
begin
  p := @target;
end;
var val1, val2: Integer;
    ptr: PInteger;
begin
  val1 := 10;
  val2 := 99;
  ptr := @val1;
  ReassignPointer(ptr, val2);
  WriteLn(ptr^);
end.
"#,
    );
    assert_eq!(out, vec!["99"]);
}

#[test]
fn test_const_string_param_no_side_effects() {
    let out = run_pascal(
        r#"
program Test;
function RepeatTwice(const s: String): String;
begin
  Result := s + s;
end;
var original: String;
begin
  original := 'ABC';
  WriteLn(RepeatTwice(original));
  WriteLn(original);
end.
"#,
    );
    assert_eq!(out, vec!["ABCABC", "ABC"]);
}

#[test]
fn test_var_param_boolean_toggle() {
    let out = run_pascal(
        r#"
program Test;
procedure Toggle(var flag: Boolean);
begin
  flag := not flag;
end;
var active: Boolean;
begin
  active := False;
  Toggle(active);
  WriteLn(active);
  Toggle(active);
  WriteLn(active);
end.
"#,
    );
    assert_eq!(out, vec!["TRUE", "FALSE"]);
}

#[test]
fn test_var_param_real_accumulation() {
    let out = run_pascal(
        r#"
program Test;
procedure Accumulate(var total: Real; amount: Real);
begin
  total := total + amount;
end;
var sum: Real;
begin
  sum := 10.5;
  Accumulate(sum, 4.5);
  WriteLn(sum);
end.
"#,
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn test_out_param_boolean_success_flag() {
    let out = run_pascal(
        r#"
program Test;
procedure TryParseInt(const s: String; out value: Integer; out success: Boolean);
begin
  if s = '123' then
  begin
    value := 123;
    success := True;
  end
  else
  begin
    value := 0;
    success := False;
  end;
end;
var val: Integer;
    ok: Boolean;
begin
  TryParseInt('123', val, ok);
  WriteLn(ok);
  WriteLn(val);
end.
"#,
    );
    assert_eq!(out, vec!["TRUE", "123"]);
}

#[test]
fn test_var_param_in_method_call() {
    let out = run_pascal(
        r#"
program Test;
type TCalculator = class
  public procedure DoubleVal(var v: Integer);
end;
procedure TCalculator.DoubleVal(var v: Integer);
begin
  v := v * 2;
end;
var calc: TCalculator;
    num: Integer;
begin
  calc := TCalculator.Create;
  num := 21;
  calc.DoubleVal(num);
  WriteLn(num);
  calc.Free;
end.
"#,
    );
    assert_eq!(out, vec!["42"]);
}
