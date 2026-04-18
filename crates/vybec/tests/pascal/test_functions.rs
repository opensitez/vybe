use super::helpers;
use helpers::run;

// Procedures
#[test] fn proc_no_params() {
    assert_eq!(run("program T; procedure Hello; begin WriteLn('hi'); end; begin Hello; end."), &["hi"]);
}
#[test] fn proc_one_param() {
    assert_eq!(run("program T; procedure Greet(name: String); begin WriteLn('Hello ' + name); end; begin Greet('World'); end."), &["Hello World"]);
}
#[test] fn proc_two_params() {
    assert_eq!(run("program T; procedure Show(a, b: Integer); begin WriteLn(a + b); end; begin Show(3, 4); end."), &["7"]);
}
#[test] fn proc_multiple_calls() {
    assert_eq!(run("program T; procedure P(x: Integer); begin WriteLn(x); end; begin P(1); P(2); P(3); end."), &["1","2","3"]);
}
#[test] fn proc_with_local_var() {
    assert_eq!(run(r#"program T;
procedure ShowDouble(x: Integer);
var r: Integer;
begin r := x * 2; WriteLn(r); end;
begin ShowDouble(5); end."#), &["10"]);
}
#[test] fn proc_modifies_via_output() {
    assert_eq!(run(r#"program T;
procedure PrintSum(a, b: Integer);
begin WriteLn(a + b); end;
begin PrintSum(10, 20); PrintSum(30, 40); end."#), &["30", "70"]);
}

// Functions
#[test] fn func_add() {
    assert_eq!(run("program T; function Add(a, b: Integer): Integer; begin Result := a + b; end; begin WriteLn(Add(3, 4)); end."), &["7"]);
}
#[test] fn func_no_params() {
    assert_eq!(run("program T; function GetFortyTwo: Integer; begin Result := 42; end; begin WriteLn(GetFortyTwo()); end."), &["42"]);
}
#[test] fn func_recursive_fact() {
    assert_eq!(run(r#"program T;
function Fact(n: Integer): Integer;
begin if n <= 1 then Result := 1 else Result := n * Fact(n - 1); end;
begin WriteLn(Fact(5)); end."#), &["120"]);
}
#[test] fn func_recursive_fib() {
    assert_eq!(run(r#"program T;
function Fib(n: Integer): Integer;
begin if n <= 1 then Result := n else Result := Fib(n - 1) + Fib(n - 2); end;
begin WriteLn(Fib(10)); end."#), &["55"]);
}
#[test] fn func_nested() {
    assert_eq!(run(r#"program T;
function Outer(x: Integer): Integer;
  function Inner(y: Integer): Integer; begin Result := y * 2; end;
begin Result := Inner(x) + 1; end;
begin WriteLn(Outer(5)); end."#), &["11"]);
}
#[test] fn func_multiple_params() {
    assert_eq!(run("program T; function Sum(a, b, c: Integer): Integer; begin Result := a + b + c; end; begin WriteLn(Sum(1, 2, 3)); end."), &["6"]);
}
#[test] fn func_string_result() {
    assert_eq!(run("program T; function Greet(name: String): String; begin Result := 'Hello ' + name; end; begin WriteLn(Greet('World')); end."), &["Hello World"]);
}
#[test] fn func_exit_early() {
    assert_eq!(run(r#"program T;
function Check(x: Integer): Integer;
begin if x > 10 then begin Result := 99; Exit; end; Result := x; end;
begin WriteLn(Check(5)); WriteLn(Check(20)); end."#), &["5", "99"]);
}
#[test] fn func_as_arg() {
    assert_eq!(run(r#"program T;
function MyDouble(x: Integer): Integer; begin Result := x * 2; end;
function MyInc(x: Integer): Integer; begin Result := x + 1; end;
begin WriteLn(MyDouble(MyInc(5))); end."#), &["12"]);
}
#[test] fn func_call_in_expr() {
    assert_eq!(run(r#"program T;
function Square(x: Integer): Integer; begin Result := x * x; end;
begin WriteLn(Square(3) + Square(4)); end."#), &["25"]);
}
#[test] fn func_assign_by_name() {
    // Pascal allows assigning to function name instead of Result
    assert_eq!(run(r#"program T;
function Triple(x: Integer): Integer;
begin Triple := x * 3; end;
begin WriteLn(Triple(7)); end."#), &["21"]);
}
#[test] fn func_bool_result() {
    assert_eq!(run(r#"program T;
function IsEven(n: Integer): Boolean;
begin Result := (n mod 2) = 0; end;
begin WriteLn(IsEven(4)); WriteLn(IsEven(7)); end."#), &["true", "false"]);
}
