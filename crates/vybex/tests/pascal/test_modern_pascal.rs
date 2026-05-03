/// Tests for modern Object Pascal features: lambdas, type casts, hex literals,
/// char literals, exit with value, nested functions, forward declarations,
/// string escapes, program/unit headings, uses clauses, parameter defaults.

use super::helpers::run_pascal;

// ===================================================================
// LAMBDAS / ANONYMOUS FUNCTIONS
// ===================================================================

#[test] fn lambda_assign_and_call() {
    assert_eq!(run_pascal(r#"program T;
var f: function(x: Integer): Integer;
begin
  f := function(x: Integer): Integer begin Result := x * 2; end;
  WriteLn(f(5));
end."#), &["10"]);
}

#[test] fn lambda_procedure_no_params() {
    assert_eq!(run_pascal(r#"program T;
var p: procedure;
begin
  p := procedure begin WriteLn('hello from lambda'); end;
  p;
end."#), &["hello from lambda"]);
}

#[test] fn lambda_as_argument() {
    assert_eq!(run_pascal(r#"program T;
procedure Apply(f: function(x: Integer): Integer; val: Integer);
begin
  WriteLn(f(val));
end;
begin
  Apply(function(x: Integer): Integer begin Result := x + 100; end, 42);
end."#), &["142"]);
}

#[test] fn lambda_captures_local() {
    assert_eq!(run_pascal(r#"program T;
var factor: Integer;
var f: function(x: Integer): Integer;
begin
  factor := 3;
  f := function(x: Integer): Integer begin Result := x * factor; end;
  WriteLn(f(10));
end."#), &["30"]);
}

// ===================================================================
// TYPE CASTS
// ===================================================================

#[test] fn cast_integer() {
    assert_eq!(run_pascal("program T; begin WriteLn(Integer(3.7)); end."), &["3"]);
}

#[test] fn cast_string() {
    assert_eq!(run_pascal("program T; begin WriteLn(String(42)); end."), &["42"]);
}

// ===================================================================
// HEX LITERALS
// ===================================================================

#[test] fn hex_literal_basic() {
    assert_eq!(run_pascal("program T; begin WriteLn($FF); end."), &["255"]);
}

#[test] fn hex_literal_zero() {
    assert_eq!(run_pascal("program T; begin WriteLn($0); end."), &["0"]);
}

#[test] fn hex_literal_arithmetic() {
    assert_eq!(run_pascal("program T; begin WriteLn($10 + $20); end."), &["48"]);
}

// ===================================================================
// CHAR LITERALS (#NNN)
// ===================================================================

#[test] fn char_literal_a() {
    assert_eq!(run_pascal("program T; begin WriteLn(#65); end."), &["A"]);
}

#[test] fn char_literal_space() {
    assert_eq!(run_pascal("program T; begin WriteLn('hello' + #32 + 'world'); end."), &["hello world"]);
}

// ===================================================================
// STRING ESCAPES
// ===================================================================

#[test] fn string_double_quote() {
    assert_eq!(run_pascal("program T; begin WriteLn('it''s'); end."), &["it's"]);
}

#[test] fn string_multiple_escapes() {
    assert_eq!(run_pascal("program T; begin WriteLn('he said ''hi'' to me'); end."), &["he said 'hi' to me"]);
}

// ===================================================================
// EXIT WITH VALUE
// ===================================================================

#[test] fn exit_with_value() {
    assert_eq!(run_pascal(r#"program T;
function Clamp(x, lo, hi: Integer): Integer;
begin
  if x < lo then Exit(lo);
  if x > hi then Exit(hi);
  Result := x;
end;
begin
  WriteLn(Clamp(-5, 0, 100));
  WriteLn(Clamp(50, 0, 100));
  WriteLn(Clamp(200, 0, 100));
end."#), &["0", "50", "100"]);
}

#[test] fn exit_no_value() {
    assert_eq!(run_pascal(r#"program T;
procedure PrintPositive(x: Integer);
begin
  if x <= 0 then Exit;
  WriteLn(x);
end;
begin
  PrintPositive(-1);
  PrintPositive(5);
  PrintPositive(0);
  PrintPositive(10);
end."#), &["5", "10"]);
}

// ===================================================================
// NESTED FUNCTIONS
// ===================================================================

#[test] fn nested_function() {
    assert_eq!(run_pascal(r#"program T;
function Outer(x: Integer): Integer;
  function Inner(y: Integer): Integer;
  begin Result := y * 2; end;
begin
  Result := Inner(x) + 1;
end;
begin
  WriteLn(Outer(5));
end."#), &["11"]);
}

#[test] fn nested_procedure() {
    assert_eq!(run_pascal(r#"program T;
procedure DoWork;
  procedure Helper;
  begin WriteLn('helper called'); end;
begin
  Helper;
end;
begin
  DoWork;
end."#), &["helper called"]);
}

// ===================================================================
// FORWARD DECLARATIONS
// ===================================================================

#[test] fn forward_declaration() {
    assert_eq!(run_pascal(r#"program T;
procedure B(n: Integer); forward;

procedure A(n: Integer);
begin
  if n > 0 then begin WriteLn('A' + IntToStr(n)); B(n - 1); end;
end;

procedure B(n: Integer);
begin
  if n > 0 then begin WriteLn('B' + IntToStr(n)); A(n - 1); end;
end;

begin
  A(3);
end."#), &["A3", "B2", "A1"]);
}

// ===================================================================
// PROGRAM / UNIT HEADING
// ===================================================================

#[test] fn program_heading() {
    assert_eq!(run_pascal("program MyApp; begin WriteLn('ok'); end."), &["ok"]);
}

// ===================================================================
// PARAMETER DEFAULTS
// ===================================================================

#[test] fn param_default_value() {
    assert_eq!(run_pascal(r#"program T;
function Greet(name: String = 'World'): String;
begin
  Result := 'Hello, ' + name + '!';
end;
begin
  WriteLn(Greet('Alice'));
  WriteLn(Greet);
end."#), &["Hello, Alice!", "Hello, World!"]);
}

// ===================================================================
// MULTIPLE PARAMS SAME TYPE (a, b, c: Integer)
// ===================================================================

#[test] fn multi_params_same_type() {
    assert_eq!(run_pascal(r#"program T;
function Sum(a, b, c: Integer): Integer;
begin Result := a + b + c; end;
begin
  WriteLn(Sum(10, 20, 30));
end."#), &["60"]);
}

// ===================================================================
// MULTIPLE VARS SAME TYPE
// ===================================================================

#[test] fn multi_vars_same_type() {
    assert_eq!(run_pascal(r#"program T;
var a, b, c: Integer;
begin
  a := 1; b := 2; c := 3;
  WriteLn(a + b + c);
end."#), &["6"]);
}

// ===================================================================
// RESULT ASSIGNMENT PATTERNS
// ===================================================================

#[test] fn result_assign_in_function() {
    assert_eq!(run_pascal(r#"program T;
function Square(x: Integer): Integer;
begin
  Result := x * x;
end;
begin WriteLn(Square(7)); end."#), &["49"]);
}

#[test] fn result_assign_conditional() {
    assert_eq!(run_pascal(r#"program T;
function AbsVal(x: Integer): Integer;
begin
  if x < 0 then Result := -x
  else Result := x;
end;
begin
  WriteLn(AbsVal(-5));
  WriteLn(AbsVal(3));
end."#), &["5", "3"]);
}

// ===================================================================
// PROCEDURE CALL WITHOUT PARENS
// ===================================================================

#[test] fn procedure_call_no_parens() {
    assert_eq!(run_pascal(r#"program T;
procedure SayHi;
begin WriteLn('hi'); end;
begin
  SayHi;
end."#), &["hi"]);
}

// ===================================================================
// FREEANDNIL REWRITE
// ===================================================================

#[test] fn freeandnil_clears_ref() {
    assert_eq!(run_pascal(r#"program T;
type TFoo = class public constructor Create; end;
constructor TFoo.Create; begin end;
var f: TFoo;
begin
  f := TFoo.Create;
  if Assigned(f) then WriteLn('assigned');
  FreeAndNil(f);
  if not Assigned(f) then WriteLn('freed');
end."#), &["assigned", "freed"]);
}

// ===================================================================
// CONST WITH TYPE HINT
// ===================================================================

#[test] fn const_with_type() {
    assert_eq!(run_pascal(r#"program T;
const MaxSize: Integer = 100;
const Greeting: String = 'Hello';
begin
  WriteLn(MaxSize);
  WriteLn(Greeting);
end."#), &["100", "Hello"]);
}

// ===================================================================
// OPERATOR PRECEDENCE
// ===================================================================

#[test] fn precedence_mul_before_add() {
    assert_eq!(run_pascal("program T; begin WriteLn(2 + 3 * 4); end."), &["14"]);
}

#[test] fn precedence_div_mod() {
    assert_eq!(run_pascal("program T; begin WriteLn(10 div 3); WriteLn(10 mod 3); end."), &["3", "1"]);
}

#[test] fn precedence_and_or() {
    assert_eq!(run_pascal("program T; begin if true or false then WriteLn('yes'); end."), &["yes"]);
}

#[test] fn precedence_not() {
    assert_eq!(run_pascal("program T; begin if not false then WriteLn('not false'); end."), &["not false"]);
}

#[test] fn precedence_parens_override() {
    assert_eq!(run_pascal("program T; begin WriteLn((2 + 3) * 4); end."), &["20"]);
}

// ===================================================================
// SHL / SHR OPERATORS
// ===================================================================

#[test] fn shl_operator() {
    assert_eq!(run_pascal("program T; begin WriteLn(1 shl 3); end."), &["8"]);
}

#[test] fn shr_operator() {
    assert_eq!(run_pascal("program T; begin WriteLn(16 shr 2); end."), &["4"]);
}

// ===================================================================
// XOR OPERATOR
// ===================================================================

#[test] fn xor_boolean() {
    assert_eq!(run_pascal("program T; begin if true xor false then WriteLn('xor works'); end."), &["xor works"]);
}
