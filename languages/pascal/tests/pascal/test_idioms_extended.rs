/// Common Object Pascal idioms — extended patterns beyond test_idioms.rs.
use super::helpers::run_pascal;

#[test]
fn idiom_break_from_inner_loop() {
    assert_eq!(
        run_pascal(
            r#"program T;
var i, j: Integer;
begin
  for i := 1 to 3 do
    for j := 1 to 3 do
      if (i = 2) and (j = 2) then Break;
  WriteLn(i); WriteLn(j);
end."#
        ),
        &["2", "2"]
    );
}

#[test]
fn idiom_continue_skip_even() {
    assert_eq!(
        run_pascal(
            r#"program T;
var i, s: Integer;
begin
  s := 0;
  for i := 1 to 6 do
  begin
    if (i mod 2) = 0 then Continue;
    s := s + i;
  end;
  WriteLn(s);
end."#
        ),
        &["9"]
    );
}

#[test]
fn idiom_repeat_until_at_least_once() {
    assert_eq!(
        run_pascal(
            r#"program T;
var n: Integer;
begin
  n := 0;
  repeat Inc(n) until n >= 3;
  WriteLn(n);
end."#
        ),
        &["3"]
    );
}

#[test]
fn idiom_case_else_branch() {
    assert_eq!(
        run_pascal(
            r#"program T;
var x: Integer;
begin
  x := 99;
  case x of
    1: WriteLn('one');
    2: WriteLn('two');
  else WriteLn('other');
  end;
end."#
        ),
        &["other"]
    );
}

#[test]
fn idiom_set_include_exclude() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TD = (A, B, C);
var s: set of TD;
begin
  s := [A, B];
  Include(s, C);
  Exclude(s, A);
  if C in s then WriteLn('c') else WriteLn('no');
  if A in s then WriteLn('a');
end."#
        ),
        &["c"]
    );
}

#[test]
fn idiom_inc_dec_on_index() {
    assert_eq!(
        run_pascal(
            r#"program T;
var i: Integer;
begin
  i := 0;
  Inc(i); Inc(i, 3); Dec(i);
  WriteLn(i);
end."#
        ),
        &["3"]
    );
}

#[test]
fn idiom_swap_xor_style() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a, b: Integer;
begin
  a := 5; b := 9;
  a := a xor b; b := a xor b; a := a xor b;
  WriteLn(a); WriteLn(b);
end."#
        ),
        &["9", "5"]
    );
}

#[test]
fn idiom_min_max_functions() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a, b: Integer;
begin
  a := 7; b := 3;
  WriteLn(Min(a, b));
  WriteLn(Max(a, b));
end."#
        ),
        &["3", "7"]
    );
}

#[test]
fn idiom_abs_negative() {
    assert_eq!(
        run_pascal(
            r#"program T;
var n: Integer;
begin
  n := -42;
  WriteLn(Abs(n));
end."#
        ),
        &["42"]
    );
}

#[test]
fn idiom_odd_even_test() {
    assert_eq!(
        run_pascal(
            r#"program T;
var n: Integer;
begin
  n := 7;
  if Odd(n) then WriteLn('odd') else WriteLn('even');
end."#
        ),
        &["odd"]
    );
}

#[test]
fn idiom_string_trim_and_upper() {
    assert_eq!(
        run_pascal(
            r#"program T;
var s: String;
begin
  s := '  hello  ';
  WriteLn(UpperCase(Trim(s)));
end."#
        ),
        &["HELLO"]
    );
}

#[test]
fn idiom_copy_and_delete_substring() {
    assert_eq!(
        run_pascal(
            r#"program T;
var s: String;
begin
  s := 'abcdef';
  Delete(s, 2, 2);
  WriteLn(s);
end."#
        ),
        &["adef"]
    );
}

#[test]
fn idiom_insert_into_string() {
    assert_eq!(
        run_pascal(
            r#"program T;
var s: String;
begin
  s := 'abef';
  Insert('cd', s, 3);
  WriteLn(s);
end."#
        ),
        &["abcdef"]
    );
}

#[test]
fn idiom_pos_find_substring() {
    assert_eq!(
        run_pascal(
            r#"program T;
var s: String;
begin
  s := 'foo-bar-baz';
  WriteLn(Pos('bar', s));
end."#
        ),
        &["5"]
    );
}

#[test]
fn idiom_length_and_setlength() {
    assert_eq!(
        run_pascal(
            r#"program T;
var s: String;
begin
  s := 'abc';
  SetLength(s, 5);
  WriteLn(Length(s));
end."#
        ),
        &["5"]
    );
}

#[test]
fn idiom_array_setlength_dynamic() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a: array of Integer;
begin
  SetLength(a, 3);
  a[0] := 1; a[1] := 2; a[2] := 3;
  WriteLn(Length(a));
end."#
        ),
        &["3"]
    );
}

#[test]
fn idiom_high_low_dynamic_array() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a: array of Integer;
begin
  SetLength(a, 4);
  WriteLn(Low(a)); WriteLn(High(a));
end."#
        ),
        &["0", "3"]
    );
}

#[test]
fn idiom_default_value_record() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TR = record X, Y: Integer; end;
var r: TR;
begin
  r := Default(TR);
  WriteLn(r.X); WriteLn(r.Y);
end."#
        ),
        &["0", "0"]
    );
}

#[test]
fn idiom_assigned_on_nil_pointer() {
    assert_eq!(
        run_pascal(
            r#"program T;
var p: ^Integer;
begin
  p := nil;
  if not Assigned(p) then WriteLn('nil');
end."#
        ),
        &["nil"]
    );
}

#[test]
fn idiom_new_dispose_pointer() {
    assert_eq!(
        run_pascal(
            r#"program T;
var p: ^Integer;
begin
  New(p);
  p^ := 77;
  WriteLn(p^);
  Dispose(p);
end."#
        ),
        &["77"]
    );
}

#[test]
fn idiom_with_record_simplifies_access() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TR = record X, Y: Integer; end;
var r: TR;
begin
  r.X := 2; r.Y := 3;
  with r do WriteLn(X + Y);
end."#
        ),
        &["5"]
    );
}

#[test]
fn idiom_nested_with_blocks() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TInner = record V: Integer; end;
type TOuter = record Inner: TInner; end;
var o: TOuter;
begin
  o.Inner.V := 8;
  with o do with Inner do WriteLn(V);
end."#
        ),
        &["8"]
    );
}

#[test]
fn idiom_exit_early_from_procedure() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Check(n: Integer);
begin
  if n < 0 then begin WriteLn('bad'); Exit; end;
  WriteLn('ok');
end;
begin Check(-1); end."#
        ),
        &["bad"]
    );
}

#[test]
fn idiom_result_assignment_in_function() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Twice(n: Integer): Integer;
begin Result := n * 2; end;
begin WriteLn(Twice(11)); end."#
        ),
        &["22"]
    );
}

#[test]
fn idiom_out_var_parameter() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Split(var lo, hi: Integer);
begin lo := 1; hi := 10; end;
var a, b: Integer;
begin
  Split(a, b);
  WriteLn(a); WriteLn(b);
end."#
        ),
        &["1", "10"]
    );
}

#[test]
fn idiom_const_parameter_no_copy_cost() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Len(const s: String): Integer;
begin Result := Length(s); end;
begin WriteLn(Len('abcd')); end."#
        ),
        &["4"]
    );
}

#[test]
fn idiom_for_in_array_iteration() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a: array[0..2] of Integer;
    v, s: Integer;
begin
  a[0]:=1; a[1]:=2; a[2]:=3;
  s := 0;
  for v in a do s := s + v;
  WriteLn(s);
end."#
        ),
        &["6"]
    );
}

#[test]
fn idiom_try_finally_free_object() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TObj = class V: Integer; end;
var o: TObj;
begin
  o := TObj.Create;
  try
    o.V := 5;
    WriteLn(o.V);
  finally
    o.Free;
    WriteLn('done');
  end;
end."#
        ),
        &["5", "done"]
    );
}

#[test]
fn idiom_raise_exception_caught() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
  raise Exception.Create('fail');
  except
    on E: Exception do WriteLn('caught');
  end;
end."#
        ),
        &["caught"]
    );
}

#[test]
fn idiom_case_insensitive_compare() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  if SameText('Hello', 'hello') then WriteLn('same') else WriteLn('diff');
end."#
        ),
        &["same"]
    );
}

#[test]
fn idiom_format_build_string() {
    assert_eq!(
        run_pascal(
            r#"program T;
var s: String;
begin
  s := Format('x=%d', [42]);
  WriteLn(s);
end."#
        ),
        &["x=42"]
    );
}

#[test]
fn idiom_inttostr_and_strtoint() {
    assert_eq!(
        run_pascal(
            r#"program T;
var s: String; n: Integer;
begin
  s := IntToStr(123);
  n := StrToInt(s);
  WriteLn(n);
end."#
        ),
        &["123"]
    );
}

#[test]
fn idiom_chr_ord_roundtrip() {
    assert_eq!(
        run_pascal(
            r#"program T;
var c: Char;
begin
  c := Chr(65);
  WriteLn(Ord(c));
end."#
        ),
        &["65"]
    );
}

#[test]
fn idiom_succ_pred_enum_steps() {
    assert_eq!(
        run_pascal(
            r#"program T;
type T = (A, B, C);
var x: T;
begin
  x := B;
  WriteLn(Ord(Succ(x)));
  WriteLn(Ord(Pred(x)));
end."#
        ),
        &["2", "0"]
    );
}

#[test]
fn idiom_boolean_and_short_circuit() {
    assert_eq!(
        run_pascal(
            r#"program T;
var called: Boolean;
function Side: Boolean;
begin called := True; Result := False; end;
begin
  called := False;
  if False and Side then WriteLn('yes');
  if called then WriteLn('called') else WriteLn('skipped');
end."#
        ),
        &["skipped"]
    );
}

#[test]
fn idiom_if_then_single_statement() {
    assert_eq!(
        run_pascal(
            r#"program T;
var n: Integer;
begin
  n := 5;
  if n > 3 then WriteLn('big');
end."#
        ),
        &["big"]
    );
}

#[test]
fn idiom_if_then_else_expression_style() {
    assert_eq!(
        run_pascal(
            r#"program T;
var n: Integer; label: String;
begin
  n := 2;
  if n > 5 then label := 'high' else label := 'low';
  WriteLn(label);
end."#
        ),
        &["low"]
    );
}

#[test]
fn idiom_while_do_zero_iterations() {
    assert_eq!(
        run_pascal(
            r#"program T;
var i: Integer;
begin
  i := 0;
  while i > 0 do Inc(i);
  WriteLn(i);
end."#
        ),
        &["0"]
    );
}

#[test]
fn idiom_for_downto_countdown() {
    assert_eq!(
        run_pascal(
            r#"program T;
var i: Integer; s: String;
begin
  s := '';
  for i := 3 downto 1 do s := s + IntToStr(i);
  WriteLn(s);
end."#
        ),
        &["321"]
    );
}

#[test]
fn idiom_string_of_char_repeat() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin WriteLn(StringOfChar('*', 5)); end."#
        ),
        &["*****"]
    );
}

#[test]
fn idiom_leading_zero_pad_manual() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Pad2(n: Integer): String;
begin
  if n < 10 then Result := '0' + IntToStr(n) else Result := IntToStr(n);
end;
begin WriteLn(Pad2(7)); WriteLn(Pad2(12)); end."#
        ),
        &["07", "12"]
    );
}

#[test]
fn idiom_guard_nil_before_deref() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TObj = class procedure Ping; end;
procedure TObj.Ping; begin WriteLn('ping'); end;
var o: TObj;
begin
  o := nil;
  if Assigned(o) then o.Ping else WriteLn('skip');
end."#
        ),
        &["skip"]
    );
}

#[test]
fn idiom_initialize_dynamic_before_use() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a: array of Integer;
begin
  SetLength(a, 2);
  WriteLn(a[0] + a[1]);
end."#
        ),
        &["0"]
    );
}

#[test]
fn idiom_compare_str_result() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  if CompareStr('abc', 'abd') < 0 then WriteLn('less') else WriteLn('geq');
end."#
        ),
        &["less"]
    );
}

#[test]
fn idiom_bool_to_str_display() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin WriteLn(BoolToStr(True, True)); end."#
        ),
        &["True"]
    );
}

#[test]
fn idiom_round_bankers_style() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin WriteLn(Round(2.6)); WriteLn(Round(2.4)); end."#
        ),
        &["3", "2"]
    );
}

#[test]
fn idiom_trunc_floor_real() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin WriteLn(Trunc(3.9)); WriteLn(Trunc(-3.9)); end."#
        ),
        &["3", "-3"]
    );
}

#[test]
fn idiom_frac_extracts_fractional_part() {
    assert_eq!(
        run_pascal(
            r#"program T;
var f: Real;
begin
  f := Frac(3.75);
  WriteLn(f > 0.7);
end."#
        ),
        &["TRUE"]
    );
}

#[test]
fn idiom_swap_strings_via_temp() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a, b, t: String;
begin
  a := 'first'; b := 'second';
  t := a; a := b; b := t;
  WriteLn(a); WriteLn(b);
end."#
        ),
        &["second", "first"]
    );
}

#[test]
fn idiom_accumulator_in_while_loop() {
    assert_eq!(
        run_pascal(
            r#"program T;
var n, sum: Integer;
begin
  n := 5; sum := 0;
  while n > 0 do begin sum := sum + n; Dec(n); end;
  WriteLn(sum);
end."#
        ),
        &["15"]
    );
}
