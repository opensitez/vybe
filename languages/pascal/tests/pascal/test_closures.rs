/// Anonymous methods, closures, function references, procedure variables,
/// callback patterns — standard Delphi/FPC anonymous method syntax.
use super::helpers::run_pascal;

// ===================================================================
// ANONYMOUS PROCEDURE
// ===================================================================

#[test]
fn anon_procedure_basic() {
    assert_eq!(
        run_pascal(
            r#"program T;
var p: procedure;
begin
  p := procedure begin WriteLn('anonymous'); end;
  p();
end."#
        ),
        &["anonymous"]
    );
}

#[test]
fn anon_procedure_with_params() {
    assert_eq!(
        run_pascal(
            r#"program T;
var p: procedure(x: Integer);
begin
  p := procedure(x: Integer) begin WriteLn(x * 2); end;
  p(5);
end."#
        ),
        &["10"]
    );
}

// ===================================================================
// ANONYMOUS FUNCTION
// ===================================================================

#[test]
fn anon_function_basic() {
    assert_eq!(
        run_pascal(
            r#"program T;
var f: function(x: Integer): Integer;
begin
  f := function(x: Integer): Integer begin Result := x * x; end;
  WriteLn(f(7));
end."#
        ),
        &["49"]
    );
}

#[test]
fn anon_function_string() {
    assert_eq!(
        run_pascal(
            r#"program T;
var f: function(s: String): String;
begin
  f := function(s: String): String begin Result := '[' + s + ']'; end;
  WriteLn(f('hello'));
end."#
        ),
        &["[hello]"]
    );
}

// ===================================================================
// CLOSURES — CAPTURE OUTER VARIABLES
// ===================================================================

#[test]
fn closure_captures_variable() {
    assert_eq!(
        run_pascal(
            r#"program T;
var multiplier: Integer;
var f: function(x: Integer): Integer;
begin
  multiplier := 3;
  f := function(x: Integer): Integer begin Result := x * multiplier; end;
  WriteLn(f(10));
end."#
        ),
        &["30"]
    );
}

#[test]
fn closure_captures_string() {
    assert_eq!(
        run_pascal(
            r#"program T;
var prefix: String;
var f: function(s: String): String;
begin
  prefix := 'Hello';
  f := function(s: String): String begin Result := prefix + ' ' + s; end;
  WriteLn(f('World'));
end."#
        ),
        &["Hello World"]
    );
}

// ===================================================================
// PASSING ANONYMOUS FUNCTIONS AS ARGUMENTS
// ===================================================================

#[test]
fn pass_function_as_arg() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Apply(arr: array of Integer; f: function(x: Integer): Integer);
var i: Integer;
begin
  for i := 0 to High(arr) do
    WriteLn(f(arr[i]));
end;
begin
  Apply([1, 2, 3], function(x: Integer): Integer begin Result := x * 10; end);
end."#
        ),
        &["10", "20", "30"]
    );
}

#[test]
fn pass_procedure_as_arg() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure ForEach(arr: array of String; p: procedure(s: String));
var i: Integer;
begin
  for i := 0 to High(arr) do p(arr[i]);
end;
begin
  ForEach(['a', 'b', 'c'], procedure(s: String) begin WriteLn(UpperCase(s)); end);
end."#
        ),
        &["A", "B", "C"]
    );
}

// ===================================================================
// FUNCTION VARIABLES — NAMED FUNCTIONS
// ===================================================================

#[test]
fn function_variable_named() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Square(x: Integer): Integer;
begin Result := x * x; end;

function Cube(x: Integer): Integer;
begin Result := x * x * x; end;

var f: function(x: Integer): Integer;
begin
  f := @Square;
  WriteLn(f(4));
  f := @Cube;
  WriteLn(f(3));
end."#
        ),
        &["16", "27"]
    );
}

// ===================================================================
// CALLBACK PATTERNS
// ===================================================================

#[test]
fn callback_on_event() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TButton = class
  public
    FOnClick: procedure(sender: String);
    constructor Create;
    procedure Click;
  end;

constructor TButton.Create; begin end;
procedure TButton.Click;
begin
  if Assigned(FOnClick) then FOnClick('Button1');
end;

var btn: TButton;
begin
  btn := TButton.Create;
  btn.FOnClick := procedure(sender: String) begin WriteLn(sender + ' clicked'); end;
  btn.Click;
end."#
        ),
        &["Button1 clicked"]
    );
}

// ===================================================================
// HIGHER-ORDER FUNCTIONS
// ===================================================================

#[test]
fn map_array() {
    assert_eq!(
        run_pascal(
            r#"program T;
function MapInt(arr: array of Integer; f: function(x: Integer): Integer): array of Integer;
var i: Integer;
begin
  Result := arr;
  for i := 0 to High(Result) do
    Result[i] := f(arr[i]);
end;

var result: array of Integer;
var i: Integer;
begin
  result := MapInt([1, 2, 3, 4], function(x: Integer): Integer begin Result := x * 2; end);
  for i := 0 to High(result) do WriteLn(result[i]);
end."#
        ),
        &["2", "4", "6", "8"]
    );
}

#[test]
fn filter_pattern() {
    assert_eq!(
        run_pascal(
            r#"program T;
var nums: array of Integer;
var n: Integer;
var evens: String;
begin
  nums := [1, 2, 3, 4, 5, 6, 7, 8];
  evens := '';
  for n in nums do
    if n mod 2 = 0 then
    begin
      if Length(evens) > 0 then evens := evens + ',';
      evens := evens + IntToStr(n);
    end;
  WriteLn(evens);
end."#
        ),
        &["2,4,6,8"]
    );
}

#[test]
fn reduce_pattern() {
    assert_eq!(
        run_pascal(
            r#"program T;
var nums: array of Integer;
var n: Integer;
var total: Integer;
begin
  nums := [1, 2, 3, 4, 5];
  total := 0;
  for n in nums do total := total + n;
  WriteLn(total);
end."#
        ),
        &["15"]
    );
}

// ===================================================================
// NESTED ANONYMOUS FUNCTIONS
// ===================================================================

#[test]
fn nested_lambda() {
    assert_eq!(
        run_pascal(
            r#"program T;
var f: function(x: Integer): function(y: Integer): Integer;
var adder: function(y: Integer): Integer;
begin
  f := function(x: Integer): function(y: Integer): Integer
    begin Result := function(y: Integer): Integer begin Result := x + y; end; end;
  adder := f(10);
  WriteLn(adder(5));
end."#
        ),
        &["15"]
    );
}
