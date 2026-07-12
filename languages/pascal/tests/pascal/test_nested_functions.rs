use super::helpers::run_pascal;

#[test]
fn test_nested_proc_basic() {
    let src = r#"
program T;
procedure Outer;
  procedure Inner;
  begin
    WriteLn('inner');
  end;
begin
  Inner;
end;
begin
  Outer;
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["inner"]);
}

#[test]
fn test_nested_proc_uses_outer_var() {
    let src = r#"
program T;
procedure Outer;
var
  x: Integer;
  procedure Inner;
  begin
    WriteLn(x);
  end;
begin
  x := 99;
  Inner;
end;
begin
  Outer;
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["99"]);
}

#[test]
fn test_nested_two_levels() {
    let src = r#"
program T;
procedure A;
  procedure B;
    procedure C;
    begin
      WriteLn('c');
    end;
  begin
    C;
    WriteLn('b');
  end;
begin
  B;
  WriteLn('a');
end;
begin
  A;
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["c", "b", "a"]);
}

#[test]
fn test_nested_func_returns_value() {
    let src = r#"
program T;
function Outer: Integer;
  function Inner: Integer;
  begin
    Result := 7;
  end;
begin
  Result := Inner * 2;
end;
begin
  WriteLn(Outer);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["14"]);
}

#[test]
fn test_mutual_recursion_even_odd() {
    let src = r#"
program T;
function IsOdd(n: Integer): Boolean; forward;

function IsEven(n: Integer): Boolean;
begin
  if n = 0 then
    Result := true
  else
    Result := IsOdd(n - 1);
end;

function IsOdd(n: Integer): Boolean;
begin
  if n = 0 then
    Result := false
  else
    Result := IsEven(n - 1);
end;

begin
  if IsEven(4) then WriteLn('even') else WriteLn('odd');
  if IsOdd(3) then WriteLn('odd') else WriteLn('even');
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["even", "odd"]);
}

#[test]
fn test_nested_modifies_outer_var() {
    let src = r#"
program T;
procedure Outer;
var
  counter: Integer;
  procedure Increment;
  begin
    counter := counter + 1;
  end;
begin
  counter := 0;
  Increment;
  Increment;
  Increment;
  WriteLn(counter);
end;
begin
  Outer;
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn test_nested_recursive() {
    let src = r#"
program T;
function Fib(n: Integer): Integer;
  function FibInner(x: Integer): Integer;
  begin
    if x <= 1 then
      Result := x
    else
      Result := FibInner(x - 1) + FibInner(x - 2);
  end;
begin
  Result := FibInner(n);
end;
begin
  WriteLn(Fib(8));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["21"]);
}

#[test]
fn test_nested_multiple_siblings() {
    let src = r#"
program T;
procedure Run;
  procedure Step1;
  begin
    WriteLn('step1');
  end;
  procedure Step2;
  begin
    WriteLn('step2');
  end;
begin
  Step1;
  Step2;
end;
begin
  Run;
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["step1", "step2"]);
}

#[test]
fn test_nested_proc_with_param() {
    let src = r#"
program T;
procedure Outer(prefix: string);
  procedure Print(msg: string);
  begin
    WriteLn(prefix + msg);
  end;
begin
  Print('hello');
  Print('world');
end;
begin
  Outer('[X] ');
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["[X] hello", "[X] world"]);
}

#[test]
fn test_nested_func_inner_loop() {
    let src = r#"
program T;
function SumTo(n: Integer): Integer;
  function Add(a, b: Integer): Integer;
  begin
    Result := a + b;
  end;
var
  i, s: Integer;
begin
  s := 0;
  for i := 1 to n do
    s := Add(s, i);
  Result := s;
end;
begin
  WriteLn(SumTo(10));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["55"]);
}

#[test]
fn test_nested_proc_call_sibling() {
    let src = r#"
program T;
procedure Parent;
  procedure A;
  begin
    WriteLn('A');
  end;
  procedure B;
  begin
    A;
    WriteLn('B');
  end;
begin
  B;
end;
begin
  Parent;
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["A", "B"]);
}

#[test]
fn test_outer_var_unchanged_after_inner() {
    let src = r#"
program T;
procedure Test;
var
  x: Integer;
  procedure UseX;
  var
    x: Integer;
  begin
    x := 100;
  end;
begin
  x := 5;
  UseX;
  WriteLn(x);
end;
begin
  Test;
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn test_forward_declaration_called_before_def() {
    let src = r#"
program T;
procedure Second; forward;

procedure First;
begin
  WriteLn('first');
  Second;
end;

procedure Second;
begin
  WriteLn('second');
end;

begin
  First;
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["first", "second"]);
}

#[test]
fn test_nested_func_called_twice() {
    let src = r#"
program T;
function Double(n: Integer): Integer;
  function Square(x: Integer): Integer;
  begin
    Result := x * x;
  end;
begin
  Result := Square(n) + Square(n);
end;
begin
  WriteLn(Double(3));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["18"]);
}

#[test]
fn test_nested_proc_in_method() {
    let src = r#"
program T;
type
  TCalc = class
    function Compute(n: Integer): Integer;
  end;

function TCalc.Compute(n: Integer): Integer;
  function Half(x: Integer): Integer;
  begin
    Result := x div 2;
  end;
begin
  Result := Half(n) + Half(n);
end;

var
  c: TCalc;
begin
  c := TCalc.Create;
  WriteLn(c.Compute(10));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["10"]);
}

#[test]
fn test_three_level_nesting_return() {
    let src = r#"
program T;
function Top: Integer;
  function Mid: Integer;
    function Bot: Integer;
    begin
      Result := 1;
    end;
  begin
    Result := Bot + 1;
  end;
begin
  Result := Mid + 1;
end;
begin
  WriteLn(Top);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn test_nested_with_while_loop() {
    let src = r#"
program T;
procedure Count;
var
  n: Integer;
  procedure PrintN;
  begin
    WriteLn(n);
  end;
begin
  n := 1;
  while n <= 3 do begin
    PrintN;
    n := n + 1;
  end;
end;
begin
  Count;
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn test_mutual_recursion_countdown() {
    let src = r#"
program T;
procedure EvenStep(n: Integer); forward;

procedure OddStep(n: Integer);
begin
  if n > 0 then begin
    WriteLn('odd:' + IntToStr(n));
    EvenStep(n - 1);
  end;
end;

procedure EvenStep(n: Integer);
begin
  if n > 0 then begin
    WriteLn('even:' + IntToStr(n));
    OddStep(n - 1);
  end;
end;

begin
  EvenStep(4);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["even:4", "odd:3", "even:2", "odd:1"]);
}

#[test]
fn test_nested_proc_conditional_call() {
    let src = r#"
program T;
procedure Evaluate(x: Integer);
  procedure Positive;
  begin
    WriteLn('positive');
  end;
  procedure Negative;
  begin
    WriteLn('negative');
  end;
  procedure Zero;
  begin
    WriteLn('zero');
  end;
begin
  if x > 0 then Positive
  else if x < 0 then Negative
  else Zero;
end;
begin
  Evaluate(5);
  Evaluate(-3);
  Evaluate(0);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["positive", "negative", "zero"]);
}

#[test]
fn test_nested_accumulates_outer_array() {
    let src = r#"
program T;
procedure BuildSum;
var
  arr: array[1..5] of Integer;
  total: Integer;
  i: Integer;
  procedure AddElement(idx, val: Integer);
  begin
    arr[idx] := val;
  end;
begin
  for i := 1 to 5 do
    AddElement(i, i * 2);
  total := 0;
  for i := 1 to 5 do
    total := total + arr[i];
  WriteLn(total);
end;
begin
  BuildSum;
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["30"]);
}
