/// Nested functions that capture outer variables — extended closure patterns.
use super::helpers::run_pascal;

#[test]
fn nested_captures_outer_integer() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Outer;
var base: Integer;
  function Inner: Integer;
  begin Result := base + 10; end;
begin
  base := 5;
  WriteLn(Inner);
end;
begin Outer; end."#
        ),
        &["15"]
    );
}

#[test]
fn nested_captures_two_outer_variables() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Outer;
var a, b: Integer;
  function Sum: Integer;
  begin Result := a + b; end;
begin
  a := 7; b := 8;
  WriteLn(Sum);
end;
begin Outer; end."#
        ),
        &["15"]
    );
}

#[test]
fn nested_modifies_captured_outer_var() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Outer;
var count: Integer;
  procedure Bump;
  begin count := count + 1; end;
begin
  count := 0;
  Bump; Bump; Bump;
  WriteLn(count);
end;
begin Outer; end."#
        ),
        &["3"]
    );
}

#[test]
fn nested_captures_outer_string() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Outer;
var label: String;
  function Tag(s: String): String;
  begin Result := label + ':' + s; end;
begin
  label := 'item';
  WriteLn(Tag('42'));
end;
begin Outer; end."#
        ),
        &["item:42"]
    );
}

#[test]
fn nested_captures_outer_boolean() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Outer;
var enabled: Boolean;
  function Status: String;
  begin if enabled then Result := 'on' else Result := 'off'; end;
begin
  enabled := true;
  WriteLn(Status);
end;
begin Outer; end."#
        ),
        &["on"]
    );
}

#[test]
fn nested_captures_record_field() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TR = record X: Integer; end;
procedure Outer;
var r: TR;
  function GetX: Integer;
  begin Result := r.X; end;
begin
  r.X := 33;
  WriteLn(GetX);
end;
begin Outer; end."#
        ),
        &["33"]
    );
}

#[test]
fn nested_captures_array_element() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Outer;
var arr: array[0..2] of Integer;
  function Middle: Integer;
  begin Result := arr[1]; end;
begin
  arr[0] := 1; arr[1] := 5; arr[2] := 9;
  WriteLn(Middle);
end;
begin Outer; end."#
        ),
        &["5"]
    );
}

#[test]
fn nested_captures_param_and_outer() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Outer(offset: Integer);
var base: Integer;
  function Calc(x: Integer): Integer;
  begin Result := base + offset + x; end;
begin
  base := 100;
  WriteLn(Calc(3));
end;
begin Outer(7); end."#
        ),
        &["110"]
    );
}

#[test]
fn nested_returns_outer_sum() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Outer;
var total: Integer;
  function Add(n: Integer): Integer;
  begin total := total + n; Result := total; end;
begin
  total := 0;
  WriteLn(Add(4));
  WriteLn(Add(6));
end;
begin Outer; end."#
        ),
        &["4", "10"]
    );
}

#[test]
fn nested_called_in_loop() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Outer;
var factor: Integer;
  function Scale(x: Integer): Integer;
  begin Result := x * factor; end;
  i: Integer;
begin
  factor := 3;
  for i := 1 to 3 do WriteLn(Scale(i));
end;
begin Outer; end."#
        ),
        &["3", "6", "9"]
    );
}

#[test]
fn nested_siblings_call_each_other() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Outer;
var seed: Integer;
  function A: Integer;
  begin Result := seed + 1; end;
  function B: Integer;
  begin Result := A + 1; end;
begin
  seed := 10;
  WriteLn(B);
end;
begin Outer; end."#
        ),
        &["12"]
    );
}

#[test]
fn nested_three_levels_capture() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure L1;
var v: Integer;
  procedure L2;
    function L3: Integer;
    begin Result := v; end;
  begin WriteLn(L3); end;
begin
  v := 77;
  L2;
end;
begin L1; end."#
        ),
        &["77"]
    );
}

#[test]
fn nested_capture_before_inner_assignment() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Outer;
var x: Integer;
  function ReadX: Integer;
  begin Result := x; end;
begin
  x := 0;
  x := 21;
  WriteLn(ReadX);
end;
begin Outer; end."#
        ),
        &["21"]
    );
}

#[test]
fn nested_closure_factorial_style() {
    assert_eq!(
        run_pascal(
            r#"program T;
function MakeFact: function(n: Integer): Integer;
var
  function Fact(n: Integer): Integer;
  begin if n <= 1 then Result := 1 else Result := n * Fact(n - 1); end;
begin Result := @Fact; end;
var f: function(n: Integer): Integer;
begin
  f := MakeFact;
  WriteLn(f(5));
end."#
        ),
        &["120"]
    );
}

#[test]
fn nested_closure_accumulator_pattern() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Outer;
var sum: Integer;
  procedure Add(n: Integer);
  begin sum := sum + n; end;
  i: Integer;
begin
  sum := 0;
  for i := 1 to 5 do Add(i);
  WriteLn(sum);
end;
begin Outer; end."#
        ),
        &["15"]
    );
}

#[test]
fn nested_captures_real_variable() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Outer;
var rate: Real;
  function Apply(v: Real): Real;
  begin Result := v * rate; end;
begin
  rate := 1.5;
  WriteLn(Apply(10));
end;
begin Outer; end."#
        ),
        &["15"]
    );
}

#[test]
fn nested_captures_char_separator() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Outer;
var sep: Char;
  function Join(a, b: String): String;
  begin Result := a + sep + b; end;
begin
  sep := '-';
  WriteLn(Join('ab', 'cd'));
end;
begin Outer; end."#
        ),
        &["ab-cd"]
    );
}

#[test]
fn nested_in_class_method_captures_field() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBox = class
  FValue: Integer;
  function DoubleIt: Integer;
end;
function TBox.DoubleIt: Integer;
  function Inner: Integer;
  begin Result := FValue * 2; end;
begin Result := Inner; end;
var b: TBox;
begin
  b := TBox.Create;
  b.FValue := 11;
  WriteLn(b.DoubleIt);
  b.Free;
end."#
        ),
        &["22"]
    );
}

#[test]
fn nested_local_shadows_outer_name() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Outer;
var x: Integer;
  function Inner: Integer;
  var x: Integer;
  begin x := 99; Result := x; end;
begin
  x := 1;
  WriteLn(Inner);
  WriteLn(x);
end;
begin Outer; end."#
        ),
        &["99", "1"]
    );
}

#[test]
fn nested_captures_enum_ordinal() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TColor = (Red, Green, Blue);
procedure Outer;
var c: TColor;
  function OrdColor: Integer;
  begin Result := Ord(c); end;
begin
  c := Blue;
  WriteLn(OrdColor);
end;
begin Outer; end."#
        ),
        &["2"]
    );
}

#[test]
fn nested_captures_set_membership() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TD = (A, B, C);
procedure Outer;
var allowed: set of TD;
  function Has(x: TD): Boolean;
  begin Result := x in allowed; end;
begin
  allowed := [A, C];
  if Has(B) then WriteLn('yes') else WriteLn('no');
end;
begin Outer; end."#
        ),
        &["no"]
    );
}

#[test]
fn nested_in_nested_loop_capture() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Outer;
var scale: Integer;
  function At(i, j: Integer): Integer;
  begin Result := (i + j) * scale; end;
  i, j: Integer;
begin
  scale := 2;
  for i := 0 to 1 do
    for j := 0 to 1 do
      WriteLn(At(i, j));
end;
begin Outer; end."#
        ),
        &["0", "2", "2", "4"]
    );
}

#[test]
fn nested_recursive_with_capture() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Outer;
var limit: Integer;
  function Down(n: Integer): Integer;
  begin if n <= 0 then Result := 0 else Result := n + Down(n - 1); end;
begin
  limit := 4;
  WriteLn(Down(limit));
end;
begin Outer; end."#
        ),
        &["10"]
    );
}

#[test]
fn nested_capture_counter_increment() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Outer;
var n: Integer;
  procedure Step;
  begin n := n + 2; end;
begin
  n := 1;
  Step; Step;
  WriteLn(n);
end;
begin Outer; end."#
        ),
        &["5"]
    );
}

#[test]
fn nested_capture_string_concat_loop() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Outer;
var acc: String;
  procedure Append(s: String);
  begin acc := acc + s; end;
begin
  acc := '';
  Append('a'); Append('b'); Append('c');
  WriteLn(acc);
end;
begin Outer; end."#
        ),
        &["abc"]
    );
}

#[test]
fn nested_in_procedure_with_local_proc() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Host;
var msg: String;
  procedure Greet(name: String);
  begin msg := 'hi ' + name; end;
begin
  Greet('sam');
  WriteLn(msg);
end;
begin Host; end."#
        ),
        &["hi sam"]
    );
}

#[test]
fn nested_capture_dynamic_array_length() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Outer;
var arr: array of Integer;
  function Len: Integer;
  begin Result := Length(arr); end;
begin
  SetLength(arr, 4);
  WriteLn(Len);
end;
begin Outer; end."#
        ),
        &["4"]
    );
}

#[test]
fn nested_capture_dynamic_array_sum() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Outer;
var data: array of Integer;
  function Total: Integer;
  var i, s: Integer;
  begin
    s := 0;
    for i := 0 to High(data) do s := s + data[i];
    Result := s;
  end;
  i: Integer;
begin
  SetLength(data, 3);
  for i := 0 to 2 do data[i] := i + 1;
  WriteLn(Total);
end;
begin Outer; end."#
        ),
        &["6"]
    );
}

#[test]
fn nested_capture_pointer_deref() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Outer;
var n: Integer;
  p: ^Integer;
  function ReadP: Integer;
  begin Result := p^; end;
begin
  n := 88;
  p := @n;
  WriteLn(ReadP);
end;
begin Outer; end."#
        ),
        &["88"]
    );
}

#[test]
fn nested_in_while_loop_capture() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Outer;
var step: Integer;
  function Next(var x: Integer): Integer;
  begin x := x + step; Result := x; end;
  x: Integer;
begin
  step := 3;
  x := 0;
  while x < 10 do WriteLn(Next(x));
end;
begin Outer; end."#
        ),
        &["3", "6", "9", "12"]
    );
}

#[test]
fn nested_in_repeat_until_capture() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Outer;
var limit: Integer;
  function Done(n: Integer): Boolean;
  begin Result := n >= limit; end;
  n: Integer;
begin
  limit := 3;
  n := 0;
  repeat
    Inc(n);
  until Done(n);
  WriteLn(n);
end;
begin Outer; end."#
        ),
        &["3"]
    );
}

#[test]
fn nested_in_for_to_capture() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Outer;
var prefix: Integer;
  function LabelOf(i: Integer): Integer;
  begin Result := prefix * 10 + i; end;
  i: Integer;
begin
  prefix := 2;
  for i := 1 to 3 do WriteLn(LabelOf(i));
end;
begin Outer; end."#
        ),
        &["21", "22", "23"]
    );
}

#[test]
fn nested_in_for_downto_capture() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Outer;
var base: Integer;
  function DecLabel(i: Integer): Integer;
  begin Result := base - i; end;
  i: Integer;
begin
  base := 10;
  for i := 2 downto 0 do WriteLn(DecLabel(i));
end;
begin Outer; end."#
        ),
        &["8", "9", "10"]
    );
}

#[test]
fn nested_multiple_return_paths() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Outer;
var threshold: Integer;
  function Classify(n: Integer): String;
  begin if n >= threshold then Result := 'high' else Result := 'low'; end;
begin
  threshold := 50;
  WriteLn(Classify(60));
  WriteLn(Classify(10));
end;
begin Outer; end."#
        ),
        &["high", "low"]
    );
}

#[test]
fn nested_inner_reads_outer_after_write() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Outer;
var x: Integer;
  procedure SetAndRead;
  begin x := 40; WriteLn(x); end;
begin
  x := 1;
  SetAndRead;
  WriteLn(x);
end;
begin Outer; end."#
        ),
        &["40", "40"]
    );
}

#[test]
fn nested_outer_reads_after_inner_write() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Outer;
var flag: Boolean;
  procedure Flip;
  begin flag := not flag; end;
begin
  flag := false;
  Flip;
  if flag then WriteLn('true') else WriteLn('false');
end;
begin Outer; end."#
        ),
        &["true"]
    );
}

#[test]
fn nested_capture_with_param_override_logic() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Outer(defaultVal: Integer);
var base: Integer;
  function Resolve(override: Boolean; v: Integer): Integer;
  begin if override then Result := v else Result := base; end;
begin
  base := 5;
  WriteLn(Resolve(false, 99));
  WriteLn(Resolve(true, 99));
end;
begin Outer(0); end."#
        ),
        &["5", "99"]
    );
}

#[test]
fn nested_deep_four_levels_capture() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure L1;
var v: Integer;
  procedure L2;
    procedure L3;
      function L4: Integer;
      begin Result := v + 1; end;
    begin WriteLn(L4); end;
  begin L3; end;
begin
  v := 6;
  L2;
end;
begin L1; end."#
        ),
        &["7"]
    );
}

#[test]
fn nested_capture_const_outer_in_expression() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Outer;
const OFFSET = 5;
var base: Integer;
  function Shifted: Integer;
  begin Result := base + OFFSET; end;
begin
  base := 10;
  WriteLn(Shifted);
end;
begin Outer; end."#
        ),
        &["15"]
    );
}

#[test]
fn nested_function_as_nested_procedure_result() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Outer;
var multiplier: Integer;
  function MakeDouble: function(x: Integer): Integer;
    function Double(x: Integer): Integer;
    begin Result := x * multiplier; end;
  begin Result := @Double; end;
var f: function(x: Integer): Integer;
begin
  multiplier := 2;
  f := MakeDouble;
  WriteLn(f(9));
end;
begin Outer; end."#
        ),
        &["18"]
    );
}

#[test]
fn nested_capture_record_updated_by_inner() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TPair = record A, B: Integer; end;
procedure Outer;
var p: TPair;
  procedure SwapFields;
  var t: Integer;
  begin t := p.A; p.A := p.B; p.B := t; end;
begin
  p.A := 1; p.B := 2;
  SwapFields;
  WriteLn(p.A); WriteLn(p.B);
end;
begin Outer; end."#
        ),
        &["2", "1"]
    );
}
