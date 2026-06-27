use super::helpers::run_pascal;

// Procedures
#[test]
fn proc_no_params() {
    assert_eq!(
        run_pascal("program T; procedure Hello; begin WriteLn('hi'); end; begin Hello; end."),
        &["hi"]
    );
}
#[test]
fn proc_one_param() {
    assert_eq!(
        run_pascal(
            "program T; procedure Greet(name: String); begin WriteLn('Hello ' + name); end; begin Greet('World'); end."
        ),
        &["Hello World"]
    );
}
#[test]
fn proc_two_params() {
    assert_eq!(
        run_pascal(
            "program T; procedure Show(a, b: Integer); begin WriteLn(a + b); end; begin Show(3, 4); end."
        ),
        &["7"]
    );
}
#[test]
fn proc_multiple_calls() {
    assert_eq!(
        run_pascal(
            "program T; procedure P(x: Integer); begin WriteLn(x); end; begin P(1); P(2); P(3); end."
        ),
        &["1", "2", "3"]
    );
}
#[test]
fn proc_with_local_var() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure ShowDouble(x: Integer);
var r: Integer;
begin r := x * 2; WriteLn(r); end;
begin ShowDouble(5); end."#
        ),
        &["10"]
    );
}
#[test]
fn proc_modifies_via_output() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure PrintSum(a, b: Integer);
begin WriteLn(a + b); end;
begin PrintSum(10, 20); PrintSum(30, 40); end."#
        ),
        &["30", "70"]
    );
}

// Functions
#[test]
fn func_add() {
    assert_eq!(
        run_pascal(
            "program T; function Add(a, b: Integer): Integer; begin Result := a + b; end; begin WriteLn(Add(3, 4)); end."
        ),
        &["7"]
    );
}
#[test]
fn func_no_params() {
    assert_eq!(
        run_pascal(
            "program T; function GetFortyTwo: Integer; begin Result := 42; end; begin WriteLn(GetFortyTwo()); end."
        ),
        &["42"]
    );
}
#[test]
fn func_recursive_fact() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Fact(n: Integer): Integer;
begin if n <= 1 then Result := 1 else Result := n * Fact(n - 1); end;
begin WriteLn(Fact(5)); end."#
        ),
        &["120"]
    );
}
#[test]
fn func_recursive_fib() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Fib(n: Integer): Integer;
begin if n <= 1 then Result := n else Result := Fib(n - 1) + Fib(n - 2); end;
begin WriteLn(Fib(10)); end."#
        ),
        &["55"]
    );
}
#[test]
fn func_nested() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Outer(x: Integer): Integer;
  function Inner(y: Integer): Integer; begin Result := y * 2; end;
begin Result := Inner(x) + 1; end;
begin WriteLn(Outer(5)); end."#
        ),
        &["11"]
    );
}
#[test]
fn func_multiple_params() {
    assert_eq!(
        run_pascal(
            "program T; function Sum(a, b, c: Integer): Integer; begin Result := a + b + c; end; begin WriteLn(Sum(1, 2, 3)); end."
        ),
        &["6"]
    );
}
#[test]
fn func_string_result() {
    assert_eq!(
        run_pascal(
            "program T; function Greet(name: String): String; begin Result := 'Hello ' + name; end; begin WriteLn(Greet('World')); end."
        ),
        &["Hello World"]
    );
}
#[test]
fn func_exit_early() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Check(x: Integer): Integer;
begin if x > 10 then begin Result := 99; Exit; end; Result := x; end;
begin WriteLn(Check(5)); WriteLn(Check(20)); end."#
        ),
        &["5", "99"]
    );
}
#[test]
fn func_as_arg() {
    assert_eq!(
        run_pascal(
            r#"program T;
function MyDouble(x: Integer): Integer; begin Result := x * 2; end;
function MyInc(x: Integer): Integer; begin Result := x + 1; end;
begin WriteLn(MyDouble(MyInc(5))); end."#
        ),
        &["12"]
    );
}
#[test]
fn func_call_in_expr() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Square(x: Integer): Integer; begin Result := x * x; end;
begin WriteLn(Square(3) + Square(4)); end."#
        ),
        &["25"]
    );
}
#[test]
fn func_assign_by_name() {
    // Pascal allows assigning to function name instead of Result
    assert_eq!(
        run_pascal(
            r#"program T;
function Triple(x: Integer): Integer;
begin Triple := x * 3; end;
begin WriteLn(Triple(7)); end."#
        ),
        &["21"]
    );
}
#[test]
fn func_bool_result() {
    assert_eq!(
        run_pascal(
            r#"program T;
function IsEven(n: Integer): Boolean;
begin Result := (n mod 2) = 0; end;
begin WriteLn(IsEven(4)); WriteLn(IsEven(7)); end."#
        ),
        &["true", "false"]
    );
}

// -------------------------------------------------------------------
// from test_functions_var_parameters.rs
// -------------------------------------------------------------------
#[test]
fn var_param_increments_caller_integer() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure IncVar(var n: Integer);
begin
  n := n + 1;
end;
var x: Integer;
begin
  x := 5;
  IncVar(x);
  WriteLn(x);
end."#
        ),
        &["6"]
    );
}

#[test]
fn var_param_swap_two_integers() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Swap(var a, b: Integer);
var t: Integer;
begin
  t := a; a := b; b := t;
end;
var x, y: Integer;
begin
  x := 1; y := 9;
  Swap(x, y);
  WriteLn(x);
  WriteLn(y);
end."#
        ),
        &["9", "1"]
    );
}

#[test]
fn var_param_string_append() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure AppendSuffix(var s: String);
begin
  s := s + '!';
end;
var msg: String;
begin
  msg := 'hi';
  AppendSuffix(msg);
  WriteLn(msg);
end."#
        ),
        &["hi!"]
    );
}

#[test]
fn var_param_boolean_toggle() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Toggle(var b: Boolean);
begin
  b := not b;
end;
var flag: Boolean;
begin
  flag := True;
  Toggle(flag);
  WriteLn(flag);
end."#
        ),
        &["false"]
    );
}

#[test]
fn var_param_record_field_mutation() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBox = record Value: Integer; end;
procedure DoubleBox(var b: TBox);
begin
  b.Value := b.Value * 2;
end;
var box: TBox;
begin
  box.Value := 7;
  DoubleBox(box);
  WriteLn(box.Value);
end."#
        ),
        &["14"]
    );
}

#[test]
fn var_param_array_element_update() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure FillSlot(var arr: array of Integer; idx: Integer; val: Integer);
begin
  arr[idx] := val;
end;
var data: array[0..2] of Integer;
begin
  FillSlot(data, 1, 42);
  WriteLn(data[1]);
end."#
        ),
        &["42"]
    );
}

#[test]
fn var_param_used_as_output_only() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure ReadCount(var n: Integer);
begin
  n := 8;
end;
var c: Integer;
begin
  ReadCount(c);
  WriteLn(c);
end."#
        ),
        &["8"]
    );
}

#[test]
fn var_param_chained_calls_accumulate() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure AddThree(var n: Integer);
begin
  n := n + 3;
end;
var total: Integer;
begin
  total := 1;
  AddThree(total);
  AddThree(total);
  WriteLn(total);
end."#
        ),
        &["7"]
    );
}

#[test]
fn var_param_with_value_param_mixed() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Scale(var n: Integer; factor: Integer);
begin
  n := n * factor;
end;
var x: Integer;
begin
  x := 4;
  Scale(x, 5);
  WriteLn(x);
end."#
        ),
        &["20"]
    );
}

#[test]
fn var_param_char_uppercase_in_place() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure UpperChar(var c: Char);
begin
  if (c >= 'a') and (c <= 'z') then
    c := Chr(Ord(c) - 32);
end;
var ch: Char;
begin
  ch := 'g';
  UpperChar(ch);
  WriteLn(ch);
end."#
        ),
        &["G"]
    );
}

#[test]
fn var_param_nested_procedure_access() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Outer;
  procedure Inner(var n: Integer);
  begin
    n := n + 10;
  end;
var v: Integer;
begin
  v := 2;
  Inner(v);
  WriteLn(v);
end;
begin
  Outer;
end."#
        ),
        &["12"]
    );
}

#[test]
fn var_param_real_halve_value() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Halve(var x: Real);
begin
  x := x / 2.0;
end;
var r: Real;
begin
  r := 9.0;
  Halve(r);
  WriteLn(r:0:1);
end."#
        ),
        &["4.5"]
    );
}

// -------------------------------------------------------------------
// from test_functions_result_return.rs
// -------------------------------------------------------------------
#[test]
fn function_result_assigned_before_return() {
    assert_eq!(
        run_pascal(
            r#"program T;
function DoubleIt(n: Integer): Integer;
begin
  Result := n * 2;
end;
begin
  WriteLn(DoubleIt(6));
end."#
        ),
        &["12"]
    );
}

#[test]
fn function_result_incremented_in_loop() {
    assert_eq!(
        run_pascal(
            r#"program T;
function SumTo(n: Integer): Integer;
var i: Integer;
begin
  Result := 0;
  for i := 1 to n do
    Result := Result + i;
end;
begin
  WriteLn(SumTo(5));
end."#
        ),
        &["15"]
    );
}

#[test]
fn function_exit_before_default_result() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Maybe(n: Integer): Integer;
begin
  Result := 0;
  if n < 0 then Exit;
  Result := n;
end;
begin
  WriteLn(Maybe(-1));
  WriteLn(Maybe(4));
end."#
        ),
        &["0", "4"]
    );
}

#[test]
fn function_string_result_concatenation() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Greet(name: String): String;
begin
  Result := 'Hello ' + name;
end;
begin
  WriteLn(Greet('Ada'));
end."#
        ),
        &["Hello Ada"]
    );
}

#[test]
fn function_boolean_result_comparison() {
    assert_eq!(
        run_pascal(
            r#"program T;
function IsEven(n: Integer): Boolean;
begin
  Result := (n mod 2) = 0;
end;
begin
  WriteLn(IsEven(8));
  WriteLn(IsEven(9));
end."#
        ),
        &["true", "false"]
    );
}

#[test]
fn function_calls_function_in_result_expression() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Twice(n: Integer): Integer; begin Result := n * 2; end;
function Quad(n: Integer): Integer; begin Result := Twice(Twice(n)); end;
begin
  WriteLn(Quad(3));
end."#
        ),
        &["12"]
    );
}

#[test]
fn procedure_with_nested_function_result() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Demo;
  function Local: Integer;
  begin
    Result := 5;
  end;
begin
  WriteLn(Local);
end;
begin
  Demo;
end."#
        ),
        &["5"]
    );
}

#[test]
fn function_result_real_division() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Average(a, b: Integer): Real;
begin
  Result := (a + b) / 2.0;
end;
begin
  WriteLn(Format('%.1f', [Average(3, 5)]));
end."#
        ),
        &["4.0"]
    );
}

#[test]
fn forward_declared_function_result() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Later(n: Integer): Integer; forward;
function Later(n: Integer): Integer;
begin
  Result := n + 1;
end;
begin
  WriteLn(Later(10));
end."#
        ),
        &["11"]
    );
}

#[test]
fn mutual_functions_odd_even() {
    assert_eq!(
        run_pascal(
            r#"program T;
function IsEven(n: Integer): Boolean; forward;
function IsOdd(n: Integer): Boolean; forward;
function IsEven(n: Integer): Boolean; begin Result := (n = 0) or IsOdd(n - 1); end;
function IsOdd(n: Integer): Boolean; begin Result := (n <> 0) and IsEven(n - 1); end;
begin
  WriteLn(IsEven(4));
  WriteLn(IsOdd(4));
end."#
        ),
        &["true", "false"]
    );
}

#[test]
fn function_const_param_not_assignable_in_body() {
    assert_eq!(
        run_pascal(
            r#"program T;
function PlusOne(const n: Integer): Integer;
begin
  Result := n + 1;
end;
begin
  WriteLn(PlusOne(8));
end."#
        ),
        &["9"]
    );
}

#[test]
fn procedure_open_array_parameter_high_bound() {
    assert_eq!(
        run_pascal(
            r#"program T;
function LastOf(const arr: array of Integer): Integer;
begin
  if Length(arr) = 0 then Result := -1 else Result := arr[High(arr)];
end;
begin
  WriteLn(LastOf([10, 20, 30]));
end."#
        ),
        &["30"]
    );
}

#[test]
fn procedure_open_array_sums_variable_length() {
    assert_eq!(
        run_pascal(
            r#"program T;
function SumOpen(const arr: array of Integer): Integer;
var i: Integer;
begin
  Result := 0;
  for i := Low(arr) to High(arr) do
    Result := Result + arr[i];
end;
begin
  WriteLn(SumOpen([1, 2, 3, 4]));
end."#
        ),
        &["10"]
    );
}

#[test]
fn nested_procedure_reads_outer_local() {
    assert_eq!(
        run_pascal(
            r#"program T;
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
end."#
        ),
        &["42"]
    );
}

#[test]
fn nested_procedure_modifies_outer_local() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Outer;
var total: Integer;
  procedure Add(n: Integer);
  begin
    total := total + n;
  end;
begin
  total := 0;
  Add(3);
  Add(4);
  WriteLn(total);
end;
begin
  Outer;
end."#
        ),
        &["7"]
    );
}

#[test]
fn function_returns_set_membership_result() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TDigit = '0'..'9';
function HasFive(const s: set of TDigit): Boolean;
begin
  Result := '5' in s;
end;
begin
  WriteLn(HasFive(['1', '5', '9']));
end."#
        ),
        &["true"]
    );
}

#[test]
fn procedure_default_parameter_optional_arg() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Greet(name: String = 'world');
begin
  WriteLn('hi ' + name);
end;
begin
  Greet;
  Greet('ada');
end."#
        ),
        &["hi world", "hi ada"]
    );
}

#[test]
fn function_pointer_variable_invocation() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TIntFunc = function(n: Integer): Integer;
function Double(n: Integer): Integer; begin Result := n * 2; end;
var fn: TIntFunc;
begin
  fn := @Double;
  WriteLn(fn(6));
end."#
        ),
        &["12"]
    );
}

#[test]
fn procedure_const_string_param_unchanged() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Show(const msg: String);
begin
  WriteLn(msg);
end;
begin
  Show('fixed');
end."#
        ),
        &["fixed"]
    );
}

#[test]
fn function_join_open_array_of_strings() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Join(const parts: array of String): String;
var i: Integer;
begin
  Result := '';
  for i := Low(parts) to High(parts) do
    Result := Result + parts[i];
end;
begin
  WriteLn(Join(['a', 'b', 'c']));
end."#
        ),
        &["abc"]
    );
}

#[test]
fn procedure_var_param_modifies_caller_variable() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Bump(var n: Integer);
begin
  n := n + 1;
end;
var x: Integer;
begin
  x := 5;
  Bump(x);
  WriteLn(x);
end."#
        ),
        &["6"]
    );
}

#[test]
fn function_result_assigned_before_exit() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Twice(n: Integer): Integer;
begin
  Result := n * 2;
end;
begin
  WriteLn(Twice(7));
end."#
        ),
        &["14"]
    );
}

#[test]
fn nested_procedure_accesses_outer_local() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Outer;
var total: Integer;
  procedure Inner;
  begin
    total := total + 3;
  end;
begin
  total := 1;
  Inner;
  WriteLn(total);
end;
begin
  Outer;
end."#
        ),
        &["4"]
    );
}

#[test]
fn forward_declaration_allows_mutual_calls() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure B(n: Integer); forward;
procedure A(n: Integer);
begin
  if n > 0 then B(n - 1);
end;
procedure B(n: Integer);
begin
  if n = 0 then WriteLn('zero');
end;
begin
  A(0);
end."#
        ),
        &["zero"]
    );
}

#[test]
fn default_parameter_value_used_when_omitted() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Greet(const name: String = 'world');
begin
  WriteLn(name);
end;
begin
  Greet;
end."#
        ),
        &["world"]
    );
}


