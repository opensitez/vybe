use super::helpers::run_pascal;

// If/Then/Else
#[test]
fn if_true() {
    assert_eq!(
        run_pascal("program T; begin if true then WriteLn('y'); end."),
        &["y"]
    );
}
#[test]
fn if_false() {
    assert_eq!(
        run_pascal("program T; begin if false then WriteLn('y'); end."),
        &[] as &[&str]
    );
}
#[test]
fn if_else_true() {
    assert_eq!(
        run_pascal("program T; begin if true then WriteLn('y') else WriteLn('n'); end."),
        &["y"]
    );
}
#[test]
fn if_else_false() {
    assert_eq!(
        run_pascal("program T; begin if false then WriteLn('y') else WriteLn('n'); end."),
        &["n"]
    );
}
#[test]
fn if_comparison() {
    assert_eq!(
        run_pascal(
            "program T; var x: Integer; begin x := 5; if x > 3 then WriteLn('y') else WriteLn('n'); end."
        ),
        &["y"]
    );
}
#[test]
fn if_nested() {
    assert_eq!(
        run_pascal(
            r#"program T; var x: Integer; begin x := 10;
      if x > 5 then if x > 8 then WriteLn('big') else WriteLn('med') else WriteLn('small'); end."#
        ),
        &["big"]
    );
}
#[test]
fn if_chained() {
    assert_eq!(
        run_pascal(
            r#"program T; var x: Integer; begin x := 2;
      if x = 1 then WriteLn('one')
      else if x = 2 then WriteLn('two')
      else WriteLn('other'); end."#
        ),
        &["two"]
    );
}
#[test]
fn if_block() {
    assert_eq!(
        run_pascal("program T; begin if true then begin WriteLn('a'); WriteLn('b'); end; end."),
        &["a", "b"]
    );
}

// For loops
#[test]
fn for_up() {
    assert_eq!(
        run_pascal("program T; var i: Integer; begin for i := 1 to 5 do WriteLn(i); end."),
        &["1", "2", "3", "4", "5"]
    );
}
#[test]
fn for_down() {
    assert_eq!(
        run_pascal("program T; var i: Integer; begin for i := 5 downto 1 do WriteLn(i); end."),
        &["5", "4", "3", "2", "1"]
    );
}
#[test]
fn for_single() {
    assert_eq!(
        run_pascal("program T; var i: Integer; begin for i := 1 to 1 do WriteLn(i); end."),
        &["1"]
    );
}
#[test]
fn for_block() {
    assert_eq!(
        run_pascal(
            "program T; var i, s: Integer; begin s := 0; for i := 1 to 5 do begin s := s + i; end; WriteLn(s); end."
        ),
        &["15"]
    );
}
#[test]
fn for_nested() {
    assert_eq!(
        run_pascal(
            "program T; var i, j: Integer; begin for i := 1 to 2 do for j := 1 to 2 do WriteLn(i * 10 + j); end."
        ),
        &["11", "12", "21", "22"]
    );
}
#[test]
fn for_zero_iterations() {
    assert_eq!(
        run_pascal(
            "program T; var i: Integer; begin for i := 5 to 3 do WriteLn('x'); WriteLn('done'); end."
        ),
        &["done"]
    );
}

// While loops
#[test]
fn while_basic() {
    assert_eq!(
        run_pascal(
            "program T; var i: Integer; begin i := 0; while i < 3 do begin WriteLn(i); i := i + 1; end; end."
        ),
        &["0", "1", "2"]
    );
}
#[test]
fn while_false() {
    assert_eq!(
        run_pascal("program T; begin while false do WriteLn('x'); end."),
        &[] as &[&str]
    );
}
#[test]
fn while_countdown() {
    assert_eq!(
        run_pascal(
            "program T; var i: Integer; begin i := 3; while i > 0 do begin WriteLn(i); i := i - 1; end; end."
        ),
        &["3", "2", "1"]
    );
}

// Repeat/Until
#[test]
fn repeat_basic() {
    assert_eq!(
        run_pascal(
            "program T; var i: Integer; begin i := 1; repeat WriteLn(i); i := i + 1; until i > 3; end."
        ),
        &["1", "2", "3"]
    );
}
#[test]
fn repeat_once() {
    assert_eq!(
        run_pascal("program T; begin repeat WriteLn('once'); until true; end."),
        &["once"]
    );
}

// Break
#[test]
fn break_for() {
    assert_eq!(
        run_pascal(
            "program T; var i: Integer; begin for i := 1 to 10 do begin if i > 3 then Break; WriteLn(i); end; end."
        ),
        &["1", "2", "3"]
    );
}
#[test]
fn break_while() {
    assert_eq!(
        run_pascal(
            "program T; var i: Integer; begin i := 0; while true do begin i := i + 1; if i > 2 then Break; WriteLn(i); end; end."
        ),
        &["1", "2"]
    );
}

// Case
#[test]
fn case_basic() {
    assert_eq!(
        run_pascal(
            "program T; var x: Integer; begin x := 2; case x of 1: WriteLn('one'); 2: WriteLn('two'); 3: WriteLn('three'); end; end."
        ),
        &["two"]
    );
}
#[test]
fn case_else() {
    assert_eq!(
        run_pascal(
            "program T; var x: Integer; begin x := 5; case x of 1: WriteLn('one'); 2: WriteLn('two'); else WriteLn('other'); end; end."
        ),
        &["other"]
    );
}
#[test]
fn case_first() {
    assert_eq!(
        run_pascal(
            "program T; var x: Integer; begin x := 1; case x of 1: WriteLn('one'); 2: WriteLn('two'); 3: WriteLn('three'); end; end."
        ),
        &["one"]
    );
}
#[test]
fn case_last() {
    assert_eq!(
        run_pascal(
            "program T; var x: Integer; begin x := 3; case x of 1: WriteLn('one'); 2: WriteLn('two'); 3: WriteLn('three'); end; end."
        ),
        &["three"]
    );
}

// -------------------------------------------------------------------
// from test_control_flow_for_bounds.rs
// -------------------------------------------------------------------
#[test]
fn for_to_negative_start_through_zero() {
    assert_eq!(
        run_pascal("program T; var i: Integer; begin for i := -2 to 0 do WriteLn(i); end."),
        &["-2", "-1", "0"]
    );
}

#[test]
fn for_downto_zero_through_negative() {
    assert_eq!(
        run_pascal("program T; var i: Integer; begin for i := 0 downto -2 do WriteLn(i); end."),
        &["0", "-1", "-2"]
    );
}

#[test]
fn for_to_skips_when_start_above_end() {
    assert_eq!(
        run_pascal(
            "program T; var i: Integer; begin for i := 5 to 2 do WriteLn(i); WriteLn('done'); end."
        ),
        &["done"]
    );
}

#[test]
fn for_downto_skips_when_start_below_end() {
    assert_eq!(
        run_pascal(
            "program T; var i: Integer; begin for i := 1 downto 4 do WriteLn(i); WriteLn('skip'); end."
        ),
        &["skip"]
    );
}

#[test]
fn for_single_iteration_when_bounds_equal() {
    assert_eq!(
        run_pascal("program T; var i: Integer; begin for i := 7 to 7 do WriteLn(i); end."),
        &["7"]
    );
}

#[test]
fn for_downto_single_iteration_when_bounds_equal() {
    assert_eq!(
        run_pascal("program T; var i: Integer; begin for i := 3 downto 3 do WriteLn(i); end."),
        &["3"]
    );
}

#[test]
fn for_nested_row_column_indices() {
    assert_eq!(
        run_pascal(
            r#"program T; var r, c: Integer; begin
  for r := 1 to 2 do
    for c := 1 to 2 do
      WriteLn(r * 10 + c);
end."#
        ),
        &["11", "12", "21", "22"]
    );
}

#[test]
fn for_break_exits_before_last_value() {
    assert_eq!(
        run_pascal(
            r#"program T; var i: Integer; begin
  for i := 1 to 6 do begin
    if i = 4 then Break;
    WriteLn(i);
  end;
end."#
        ),
        &["1", "2", "3"]
    );
}

#[test]
fn for_continue_skips_selected_values() {
    assert_eq!(
        run_pascal(
            r#"program T; var i: Integer; begin
  for i := 1 to 4 do begin
    if i = 2 then Continue;
    WriteLn(i);
  end;
end."#
        ),
        &["1", "3", "4"]
    );
}

#[test]
fn for_loop_variable_persists_after_loop() {
    assert_eq!(
        run_pascal("program T; var i: Integer; begin for i := 1 to 3 do ; WriteLn(i); end."),
        &["4"]
    );
}

#[test]
fn for_downto_loop_variable_after_exit_value() {
    assert_eq!(
        run_pascal("program T; var i: Integer; begin for i := 5 downto 1 do ; WriteLn(i); end."),
        &["0"]
    );
}

#[test]
fn for_with_begin_block_accumulates_sum() {
    assert_eq!(
        run_pascal(
            r#"program T; var i, s: Integer; begin
  s := 0;
  for i := 1 to 5 do begin
    s := s + i;
  end;
  WriteLn(s);
end."#
        ),
        &["15"]
    );
}

#[test]
fn for_to_with_char_counter_display() {
    assert_eq!(
        run_pascal(
            r#"program T; var c: Char; begin
  for c := 'a' to 'c' do WriteLn(c);
end."#
        ),
        &["a", "b", "c"]
    );
}

#[test]
fn for_nested_break_inner_only() {
    assert_eq!(
        run_pascal(
            r#"program T; var i, j: Integer; begin
  for i := 1 to 2 do
    for j := 1 to 3 do begin
      if j = 2 then Break;
      WriteLn(i * 10 + j);
    end;
end."#
        ),
        &["11", "21"]
    );
}

#[test]
fn for_labelled_inner_loop_continues_outer() {
    assert_eq!(
        run_pascal(
            r#"program T; var i, j, s: Integer; begin
  s := 0;
  for i := 1 to 2 do begin
    for j := 1 to 2 do begin
      if (i = 2) and (j = 1) then Continue;
      s := s + 1;
    end;
  end;
  WriteLn(s);
end."#
        ),
        &["3"]
    );
}

#[test]
fn for_exit_procedure_from_inside_for() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Scan;
var i: Integer;
begin
  for i := 1 to 10 do begin
    if i = 3 then Exit;
  end;
  WriteLn('never');
end;
begin
  Scan;
  WriteLn('after');
end."#
        ),
        &["after"]
    );
}

#[test]
fn for_with_enum_range() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TDay = (Mon, Tue, Wed);
var d: TDay;
begin
  for d := Mon to Wed do WriteLn(Ord(d));
end."#
        ),
        &["0", "1", "2"]
    );
}

#[test]
fn for_inner_modifies_outer_accumulator() {
    assert_eq!(
        run_pascal(
            r#"program T; var i, j, p: Integer; begin
  p := 1;
  for i := 1 to 3 do
    for j := 1 to 2 do
      p := p + 1;
  WriteLn(p);
end."#
        ),
        &["7"]
    );
}

#[test]
fn for_downto_counts_removals_from_stack() {
    assert_eq!(
        run_pascal(
            r#"program T; var i, n: Integer; begin
  n := 5;
  for i := n downto 1 do
    n := n - 1;
  WriteLn(n);
end."#
        ),
        &["0"]
    );
}

#[test]
fn for_empty_body_still_advances_counter() {
    assert_eq!(
        run_pascal("program T; var i: Integer; begin for i := 1 to 100 do ; WriteLn(i); end."),
        &["101"]
    );
}

// -------------------------------------------------------------------
// from test_control_flow_repeat_until.rs
// -------------------------------------------------------------------
#[test]
fn repeat_until_false_runs_once() {
    assert_eq!(
        run_pascal(
            r#"program T; var n: Integer; begin
  n := 0;
  repeat
    n := n + 1;
  until n >= 1;
  WriteLn(n);
end."#
        ),
        &["1"]
    );
}

#[test]
fn repeat_until_true_skips_second_body() {
    assert_eq!(
        run_pascal(
            r#"program T; var n: Integer; begin
  n := 0;
  repeat
    n := n + 1;
    WriteLn(n);
  until n >= 1;
end."#
        ),
        &["1"]
    );
}

#[test]
fn repeat_until_multiple_iterations() {
    assert_eq!(
        run_pascal(
            r#"program T; var n: Integer; begin
  n := 0;
  repeat
    n := n + 1;
  until n = 4;
  WriteLn(n);
end."#
        ),
        &["4"]
    );
}

#[test]
fn repeat_break_before_until() {
    assert_eq!(
        run_pascal(
            r#"program T; var n: Integer; begin
  n := 0;
  repeat
    n := n + 1;
    if n = 2 then Break;
  until n > 10;
  WriteLn(n);
end."#
        ),
        &["2"]
    );
}

#[test]
fn repeat_continue_skips_write() {
    assert_eq!(
        run_pascal(
            r#"program T; var n: Integer; begin
  n := 0;
  repeat
    n := n + 1;
    if n = 2 then Continue;
    WriteLn(n);
  until n >= 3;
end."#
        ),
        &["1", "3"]
    );
}

#[test]
fn while_zero_iterations_when_condition_false() {
    assert_eq!(
        run_pascal(
            r#"program T; begin
  while False do WriteLn('x');
  WriteLn('done');
end."#
        ),
        &["done"]
    );
}

#[test]
fn while_counts_down_to_zero() {
    assert_eq!(
        run_pascal(
            r#"program T; var n: Integer; begin
  n := 3;
  while n > 0 do begin
    WriteLn(n);
    n := n - 1;
  end;
end."#
        ),
        &["3", "2", "1"]
    );
}

#[test]
fn while_break_exits_before_condition_recheck() {
    assert_eq!(
        run_pascal(
            r#"program T; var n: Integer; begin
  n := 0;
  while n < 5 do begin
    n := n + 1;
    if n = 3 then Break;
  end;
  WriteLn(n);
end."#
        ),
        &["3"]
    );
}

#[test]
fn while_continue_skips_tail_of_body() {
    assert_eq!(
        run_pascal(
            r#"program T; var n: Integer; begin
  n := 0;
  while n < 3 do begin
    n := n + 1;
    if n = 2 then Continue;
    WriteLn(n);
  end;
end."#
        ),
        &["1", "3"]
    );
}

#[test]
fn repeat_nested_until_inner_condition() {
    assert_eq!(
        run_pascal(
            r#"program T; var i, j: Integer; begin
  i := 0;
  repeat
    i := i + 1;
    j := 0;
    repeat
      j := j + 1;
    until j = 2;
    WriteLn(i * 10 + j);
  until i = 2;
end."#
        ),
        &["12", "22"]
    );
}

#[test]
fn while_condition_evaluated_each_iteration() {
    assert_eq!(
        run_pascal(
            r#"program T; var n: Integer; begin
  n := 1;
  while n < 4 do begin
    WriteLn(n);
    n := n * 2;
  end;
end."#
        ),
        &["1", "2"]
    );
}

#[test]
fn repeat_until_with_boolean_flag() {
    assert_eq!(
        run_pascal(
            r#"program T; var done: Boolean; n: Integer; begin
  n := 0;
  done := False;
  repeat
    n := n + 1;
    if n = 3 then done := True;
  until done;
  WriteLn(n);
end."#
        ),
        &["3"]
    );
}

#[test]
fn while_with_string_empty_check() {
    assert_eq!(
        run_pascal(
            r#"program T; var s: String; begin
  s := 'ab';
  while Length(s) > 0 do begin
    WriteLn(s[1]);
    Delete(s, 1, 1);
  end;
end."#
        ),
        &["a", "b"]
    );
}

#[test]
fn repeat_until_char_sequence() {
    assert_eq!(
        run_pascal(
            r#"program T; var c: Char; begin
  c := 'x';
  repeat
    WriteLn(c);
    c := 'y';
  until c = 'y';
end."#
        ),
        &["x"]
    );
}

#[test]
fn while_exit_procedure_from_loop() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Work;
var n: Integer;
begin
  n := 0;
  while n < 10 do begin
    n := n + 1;
    if n = 2 then Exit;
  end;
  WriteLn('inner');
end;
begin
  Work;
  WriteLn('outer');
end."#
        ),
        &["outer"]
    );
}

// -------------------------------------------------------------------
// from test_control_flow_case_ranges.rs
// -------------------------------------------------------------------
#[test]
fn case_range_inclusive_hits_middle_value() {
    assert_eq!(
        run_pascal(
            r#"program T; var x: Integer; begin
  x := 15;
  case x of
    10..20: WriteLn('mid');
  else
    WriteLn('other');
  end;
end."#
        ),
        &["mid"]
    );
}

#[test]
fn case_multiple_labels_share_branch() {
    assert_eq!(
        run_pascal(
            r#"program T; var x: Integer; begin
  x := 2;
  case x of
    1, 2, 3: WriteLn('small');
  else
    WriteLn('big');
  end;
end."#
        ),
        &["small"]
    );
}

#[test]
fn case_char_literal_branch() {
    assert_eq!(
        run_pascal(
            r#"program T; var c: Char; begin
  c := 'z';
  case c of
    'a'..'m': WriteLn('first-half');
    'n'..'z': WriteLn('second-half');
  end;
end."#
        ),
        &["second-half"]
    );
}

#[test]
fn case_enum_dispatches_by_ordinal() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TSize = (Small, Medium, Large);
var s: TSize;
begin
  s := Medium;
  case s of
    Small: WriteLn('s');
    Medium: WriteLn('m');
    Large: WriteLn('l');
  end;
end."#
        ),
        &["m"]
    );
}

#[test]
fn case_else_when_no_label_matches() {
    assert_eq!(
        run_pascal(
            r#"program T; var x: Integer; begin
  x := 99;
  case x of
    1: WriteLn('one');
  else
    WriteLn('default');
  end;
end."#
        ),
        &["default"]
    );
}

#[test]
fn case_nested_inside_for_loop() {
    assert_eq!(
        run_pascal(
            r#"program T; var i: Integer; begin
  for i := 1 to 3 do
    case i of
      1: WriteLn('a');
      2: WriteLn('b');
      3: WriteLn('c');
    end;
end."#
        ),
        &["a", "b", "c"]
    );
}

#[test]
fn case_negative_value_range() {
    assert_eq!(
        run_pascal(
            r#"program T; var x: Integer; begin
  x := -3;
  case x of
    -5..-1: WriteLn('neg');
  else
    WriteLn('other');
  end;
end."#
        ),
        &["neg"]
    );
}

#[test]
fn case_boolean_true_branch_only() {
    assert_eq!(
        run_pascal(
            r#"program T; var b: Boolean; begin
  b := True;
  case b of
    True: WriteLn('yes');
    False: WriteLn('no');
  end;
end."#
        ),
        &["yes"]
    );
}

#[test]
fn case_function_result_as_selector() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Tag: Integer;
begin
  Result := 4;
end;
begin
  case Tag of
    4: WriteLn('four');
  else
    WriteLn('other');
  end;
end."#
        ),
        &["four"]
    );
}

#[test]
fn case_with_begin_end_block() {
    assert_eq!(
        run_pascal(
            r#"program T; var x: Integer; begin
  x := 1;
  case x of
    1: begin WriteLn('one'); WriteLn('uno'); end;
  end;
end."#
        ),
        &["one", "uno"]
    );
}

#[test]
fn if_without_else_skips_false_branch() {
    assert_eq!(
        run_pascal(r#"program T; begin if False then WriteLn('yes'); WriteLn('after'); end."#),
        &["after"]
    );
}

#[test]
fn if_and_short_circuit_skips_second_predicate() {
    assert_eq!(
        run_pascal(
            r#"program T;
var n: Integer;
begin
  n := 0;
  if False and (n = 1) then WriteLn('bad') else WriteLn('ok');
end."#
        ),
        &["ok"]
    );
}

#[test]
fn if_or_short_circuit_skips_second_predicate() {
    assert_eq!(
        run_pascal(
            r#"program T;
var n: Integer;
begin
  n := 0;
  if True or (n = 1) then WriteLn('ok') else WriteLn('bad');
end."#
        ),
        &["ok"]
    );
}

#[test]
fn if_not_inverts_boolean_expression() {
    assert_eq!(
        run_pascal(r#"program T; begin if not False then WriteLn('yes'); end."#),
        &["yes"]
    );
}

#[test]
fn if_nested_else_binds_to_nearest_if() {
    assert_eq!(
        run_pascal(
            r#"program T; var x: Integer; begin
  x := 2;
  if x = 1 then WriteLn('a')
  else if x = 2 then WriteLn('b')
  else WriteLn('c');
end."#
        ),
        &["b"]
    );
}

#[test]
fn repeat_until_condition_checked_after_body() {
    assert_eq!(
        run_pascal(
            r#"program T; var n: Integer; begin
  n := 0;
  repeat
    n := n + 1;
    WriteLn('body');
  until n >= 2;
  WriteLn(n);
end."#
        ),
        &["body", "body", "2"]
    );
}

#[test]
fn while_loop_condition_checked_before_body() {
    assert_eq!(
        run_pascal(
            r#"program T; begin
  while False do WriteLn('never');
  WriteLn('done');
end."#
        ),
        &["done"]
    );
}

#[test]
fn case_min_value_branch_selected() {
    assert_eq!(
        run_pascal(
            r#"program T; var x: Integer; begin
  x := Low(Integer);
  case x of
    Low(Integer): WriteLn('min');
  else
    WriteLn('other');
  end;
end."#
        ),
        &["min"]
    );
}

#[test]
fn for_loop_accumulates_with_explicit_counter() {
    assert_eq!(
        run_pascal(
            r#"program T; var i, sum: Integer; begin
  sum := 0;
  for i := 1 to 3 do
    sum := sum + i;
  WriteLn(sum);
end."#
        ),
        &["6"]
    );
}

#[test]
fn if_then_without_else_skips_false_branch() {
    assert_eq!(
        run_pascal(
            r#"program T; begin
  if False then WriteLn('yes');
  WriteLn('no');
end."#
        ),
        &["no"]
    );
}

#[test]
fn case_else_catches_unlisted_value() {
    assert_eq!(
        run_pascal(
            r#"program T; var c: Char; begin
  c := 'z';
  case c of
    'a': WriteLn('a');
    'b': WriteLn('b');
  else
    WriteLn('other');
  end;
end."#
        ),
        &["other"]
    );
}

#[test]
fn repeat_until_runs_at_least_once() {
    assert_eq!(
        run_pascal(
            r#"program T; begin
  repeat
    WriteLn('once');
  until True;
end."#
        ),
        &["once"]
    );
}

#[test]
fn nested_for_loops_print_row_column() {
    assert_eq!(
        run_pascal(
            r#"program T; var r, c: Integer; begin
  for r := 1 to 2 do
    for c := 1 to 2 do
      WriteLn(r * 10 + c);
end."#
        ),
        &["11", "12", "21", "22"]
    );
}

#[test]
fn break_exits_innermost_loop_only() {
    assert_eq!(
        run_pascal(
            r#"program T; var i, j: Integer; begin
  for i := 1 to 2 do begin
    for j := 1 to 3 do begin
      if j = 2 then Break;
      WriteLn(j);
    end;
    WriteLn('row');
  end;
end."#
        ),
        &["1", "row", "1", "row"]
    );
}
