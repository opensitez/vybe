/// Advanced function patterns: overload, default params, nesting, Result usage.
use super::helpers::run_pascal;

#[test]
fn overload_integer_and_string_append() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Combine(a, b: Integer): Integer; overload;
function Combine(a, b: String): String; overload;
function Combine(a, b: Integer): Integer; begin Result := a + b; end;
function Combine(a, b: String): String; begin Result := a + b; end;
begin
  WriteLn(Combine(2, 3));
  WriteLn(Combine('x', 'y'));
end."#
        ),
        &["5", "xy"]
    );
}

#[test]
fn overload_single_and_double_params() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Twice(n: Integer): Integer; overload;
function Twice(a, b: Integer): Integer; overload;
function Twice(n: Integer): Integer; begin Result := n * 2; end;
function Twice(a, b: Integer): Integer; begin Result := a + b; end;
begin
  WriteLn(Twice(4));
  WriteLn(Twice(1, 2));
end."#
        ),
        &["8", "3"]
    );
}

#[test]
fn overload_real_and_integer_abs() {
    assert_eq!(
        run_pascal(
            r#"program T;
function MyAbs(v: Integer): Integer; overload;
function MyAbs(v: Double): Double; overload;
function MyAbs(v: Integer): Integer; begin Result := Abs(v); end;
function MyAbs(v: Double): Double; begin Result := Abs(v); end;
begin
  WriteLn(MyAbs(-7));
  WriteLn(MyAbs(-2.5) > 2.0);
end."#
        ),
        &["7", "TRUE"]
    );
}

#[test]
fn default_param_integer_multiplier() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Scale(n: Integer; factor: Integer = 2): Integer;
begin Result := n * factor; end;
begin
  WriteLn(Scale(5));
  WriteLn(Scale(5, 3));
end."#
        ),
        &["10", "15"]
    );
}

#[test]
fn default_param_string_prefix() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Show(msg: String; prefix: String = '[info] ');
begin WriteLn(prefix + msg); end;
begin
  Show('ready');
  Show('done', '>> ');
end."#
        ),
        &["[info] ready", ">> done"]
    );
}

#[test]
fn two_default_params_second_omitted() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Pair(a: Integer = 1; b: Integer = 2): Integer;
begin Result := a + b; end;
begin
  WriteLn(Pair());
  WriteLn(Pair(10));
  WriteLn(Pair(10, 20));
end."#
        ),
        &["3", "12", "30"]
    );
}

#[test]
fn default_bool_flag_controls_output() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Emit(tag: String; loud: Boolean = false);
begin
  if loud then WriteLn(UpperCase(tag)) else WriteLn(tag);
end;
begin
  Emit('soft');
  Emit('loud', true);
end."#
        ),
        &["soft", "LOUD"]
    );
}

#[test]
fn nested_function_triple_depth() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Outer(x: Integer): Integer;
  function Mid(y: Integer): Integer;
    function Inner(z: Integer): Integer;
    begin Result := z + 1; end;
  begin Result := Inner(y) * 2; end;
begin Result := Mid(x) + 3; end;
begin WriteLn(Outer(4)); end."#
        ),
        &["13"]
    );
}

#[test]
fn nested_procedure_calls_sibling_inner() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Runner;
  procedure A; forward;
  procedure B;
  begin WriteLn('B'); A; end;
  procedure A;
  begin WriteLn('A'); end;
begin B; end;
begin Runner; end."#
        ),
        &["B", "A"]
    );
}

#[test]
fn nested_function_closure_over_outer_param() {
    assert_eq!(
        run_pascal(
            r#"program T;
function MakeAdder(base: Integer): Integer;
  function Add(n: Integer): Integer;
  begin Result := base + n; end;
begin Result := Add(5); end;
begin WriteLn(MakeAdder(10)); end."#
        ),
        &["15"]
    );
}

#[test]
fn result_set_incrementally_in_loop() {
    assert_eq!(
        run_pascal(
            r#"program T;
function SumTo(n: Integer): Integer;
var i: Integer;
begin
  Result := 0;
  for i := 1 to n do Result := Result + i;
end;
begin WriteLn(SumTo(5)); end."#
        ),
        &["15"]
    );
}

#[test]
fn result_assigned_in_each_branch() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Sign(n: Integer): Integer;
begin
  if n > 0 then Result := 1
  else if n < 0 then Result := -1
  else Result := 0;
end;
begin
  WriteLn(Sign(3));
  WriteLn(Sign(0));
  WriteLn(Sign(-2));
end."#
        ),
        &["1", "0", "-1"]
    );
}

#[test]
fn result_string_built_character_by_character() {
    assert_eq!(
        run_pascal(
            r#"program T;
function RepeatChar(ch: Char; n: Integer): String;
var i: Integer;
begin
  Result := '';
  for i := 1 to n do Result := Result + ch;
end;
begin WriteLn(RepeatChar('*', 4)); end."#
        ),
        &["****"]
    );
}

#[test]
fn function_returns_record_by_value() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TPair = record A, B: Integer; end;
function MakePair(x, y: Integer): TPair;
begin Result.A := x; Result.B := y; end;
var p: TPair;
begin
  p := MakePair(2, 5);
  WriteLn(p.A + p.B);
end."#
        ),
        &["7"]
    );
}

#[test]
fn function_returns_dynamic_array() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Range(n: Integer): array of Integer;
var i: Integer;
begin
  SetLength(Result, n);
  for i := 0 to n - 1 do Result[i] := i + 1;
end;
var a: array of Integer;
begin
  a := Range(3);
  WriteLn(a[2]);
end."#
        ),
        &["3"]
    );
}

#[test]
fn overload_procedure_by_param_count() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Log(msg: String); overload;
procedure Log(tag, msg: String); overload;
procedure Log(msg: String); begin WriteLn(msg); end;
procedure Log(tag, msg: String); begin WriteLn(tag + ': ' + msg); end;
begin
  Log('plain');
  Log('warn', 'slow');
end."#
        ),
        &["plain", "warn: slow"]
    );
}

#[test]
fn default_const_param_preserves_literal() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Echo(const msg: String = 'default');
begin WriteLn(msg); end;
begin Echo; end."#
        ),
        &["default"]
    );
}

#[test]
fn nested_function_shadows_outer_name() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Outer: Integer;
var v: Integer;
  function Inner: Integer;
  var v: Integer;
  begin v := 2; Result := v; end;
begin v := 1; Result := v + Inner; end;
begin WriteLn(Outer); end."#
        ),
        &["3"]
    );
}

#[test]
fn function_exit_preserves_prior_result() {
    assert_eq!(
        run_pascal(
            r#"program T;
function FirstPositive(const a: array of Integer): Integer;
var i: Integer;
begin
  Result := -1;
  for i := Low(a) to High(a) do
    if a[i] > 0 then begin Result := a[i]; Exit; end;
end;
begin WriteLn(FirstPositive([-3, 0, 7, 2])); end."#
        ),
        &["7"]
    );
}

#[test]
fn recursive_function_with_result_accumulator() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Power(base, exp: Integer): Integer;
begin
  if exp = 0 then Result := 1
  else Result := base * Power(base, exp - 1);
end;
begin WriteLn(Power(2, 8)); end."#
        ),
        &["256"]
    );
}

#[test]
fn nested_procedure_modifies_outer_result_var() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Compute: Integer;
  procedure Add(n: Integer);
  begin Result := Result + n; end;
begin
  Result := 0;
  Add(4);
  Add(6);
end;
begin WriteLn(Compute); end."#
        ),
        &["10"]
    );
}

#[test]
fn overload_char_and_string_first() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Head(c: Char): String; overload;
function Head(const s: String): Char; overload;
function Head(c: Char): String; begin Result := c; end;
function Head(const s: String): Char; begin Result := s[1]; end;
begin
  WriteLn(Head('Z'));
  WriteLn(Head('abc'));
end."#
        ),
        &["Z", "a"]
    );
}

#[test]
fn default_param_real_tolerance() {
    assert_eq!(
        run_pascal(
            r#"program T;
function NearEqual(a, b: Double; eps: Double = 0.001): Boolean;
begin Result := Abs(a - b) <= eps; end;
begin
  WriteLn(NearEqual(1.0, 1.0005));
  WriteLn(NearEqual(1.0, 1.01, 0.1));
end."#
        ),
        &["true", "true"]
    );
}

#[test]
fn function_returns_boolean_from_comparison() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Between(n, lo, hi: Integer): Boolean;
begin Result := (n >= lo) and (n <= hi); end;
begin
  WriteLn(Between(5, 1, 10));
  WriteLn(Between(11, 1, 10));
end."#
        ),
        &["true", "false"]
    );
}

#[test]
fn nested_triple_procedure_print_order() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure L1;
  procedure L2;
    procedure L3;
    begin WriteLn(3); end;
  begin WriteLn(2); L3; end;
begin WriteLn(1); L2; end;
begin L1; end."#
        ),
        &["1", "2", "3"]
    );
}

#[test]
fn result_enum_from_function() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TDir = (North, East, South, West);
function TurnRight(d: TDir): TDir;
begin Result := Succ(d); end;
begin WriteLn(Ord(TurnRight(North))); end."#
        ),
        &["1"]
    );
}

#[test]
fn overload_array_open_and_fixed() {
    assert_eq!(
        run_pascal(
            r#"program T;
function First(const a: array of Integer): Integer; overload;
function First(const a: array[0..1] of Integer): Integer; overload;
function First(const a: array of Integer): Integer; begin Result := a[0]; end;
function First(const a: array[0..1] of Integer): Integer; begin Result := a[1]; end;
var fixed: array[0..1] of Integer;
    dyn: array of Integer;
begin
  fixed[0] := 10; fixed[1] := 20;
  dyn := [5, 6];
  WriteLn(First(fixed));
  WriteLn(First(dyn));
end."#
        ),
        &["20", "5"]
    );
}

#[test]
fn default_string_param_with_expression_call() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Tag(const name: String = 'anon');
begin WriteLn('<' + name + '>'); end;
begin
  Tag;
  Tag('vybe');
end."#
        ),
        &["<anon>", "<vybe>"]
    );
}

#[test]
fn function_result_reassigned_multiple_times() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Pick(n: Integer): String;
begin
  Result := 'none';
  if n = 1 then Result := 'one';
  if n = 2 then Result := 'two';
end;
begin WriteLn(Pick(2)); end."#
        ),
        &["two"]
    );
}

#[test]
fn nested_function_returns_outer_local() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Wrapper: Integer;
var acc: Integer;
  function Inner: Integer;
  begin Result := acc; end;
begin
  acc := 42;
  Result := Inner;
end;
begin WriteLn(Wrapper); end."#
        ),
        &["42"]
    );
}

#[test]
fn overload_boolean_negators() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Flip(v: Boolean): Boolean; overload;
function Flip(v: Integer): Integer; overload;
function Flip(v: Boolean): Boolean; begin Result := not v; end;
function Flip(v: Integer): Integer; begin Result := -v; end;
begin
  WriteLn(Flip(true));
  WriteLn(Flip(9));
end."#
        ),
        &["false", "-9"]
    );
}

#[test]
fn default_param_with_explicit_zero() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Offset(n: Integer; delta: Integer = 0): Integer;
begin Result := n + delta; end;
begin
  WriteLn(Offset(5));
  WriteLn(Offset(5, 0));
end."#
        ),
        &["5", "5"]
    );
}

#[test]
fn function_returns_set_type() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TD = '0'..'9';
function Digits: set of TD;
begin Result := ['1', '3', '5']; end;
var s: set of TD;
begin
  s := Digits;
  WriteLn('3' in s);
end."#
        ),
        &["true"]
    );
}

#[test]
fn nested_mutual_recursion_functions() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Even(n: Integer): Boolean; forward;
function Odd(n: Integer): Boolean;
begin
  if n = 0 then Result := false else Result := Even(n - 1);
end;
function Even(n: Integer): Boolean;
begin
  if n = 0 then Result := true else Result := Odd(n - 1);
end;
begin
  WriteLn(Odd(3));
  WriteLn(Even(4));
end."#
        ),
        &["true", "true"]
    );
}

#[test]
fn result_real_from_integer_division() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Ratio(a, b: Integer): Double;
begin Result := a / b; end;
begin WriteLn(Ratio(7, 2) > 3.0); end."#
        ),
        &["TRUE"]
    );
}

#[test]
fn overload_procedure_no_args_and_one_arg() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Ping; overload;
procedure Ping(n: Integer); overload;
procedure Ping; begin WriteLn('ping'); end;
procedure Ping(n: Integer); begin WriteLn(n); end;
begin
  Ping;
  Ping(7);
end."#
        ),
        &["ping", "7"]
    );
}

#[test]
fn default_nested_call_uses_middle_default() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Mul(a: Integer; b: Integer = 2; c: Integer = 3): Integer;
begin Result := a * b * c; end;
begin WriteLn(Mul(2)); WriteLn(Mul(2, 5)); end."#
        ),
        &["12", "30"]
    );
}

#[test]
fn function_returns_char_from_ord() {
    assert_eq!(
        run_pascal(
            r#"program T;
function CodeChar(n: Integer): Char;
begin Result := Chr(n); end;
begin WriteLn(CodeChar(72)); end."#
        ),
        &["H"]
    );
}

#[test]
fn nested_procedure_with_local_function_call() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Driver;
  function Local: Integer;
  begin Result := 9; end;
begin WriteLn(Local); end;
begin Driver; end."#
        ),
        &["9"]
    );
}

#[test]
fn result_initialized_then_conditionally_overwritten() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Classify(n: Integer): String;
begin
  Result := 'other';
  if n mod 2 = 0 then Result := 'even';
  if n mod 2 = 1 then Result := 'odd';
end;
begin WriteLn(Classify(6)); WriteLn(Classify(5)); end."#
        ),
        &["even", "odd"]
    );
}
