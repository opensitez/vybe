use super::helpers::run_pascal;

#[test]
fn test_goto_simple_forward() {
    let src = r#"
program T;
label done;
var
  x: Integer;
begin
  x := 1;
  if x = 1 then goto done;
  WriteLn('skipped');
  done:
  WriteLn('reached');
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["reached"]);
}

#[test]
fn test_goto_loop_simulation() {
    let src = r#"
program T;
label loop_start;
var
  i: Integer;
begin
  i := 0;
  loop_start:
  if i < 5 then begin
    WriteLn(i);
    i := i + 1;
    goto loop_start;
  end;
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["0", "1", "2", "3", "4"]);
}

#[test]
fn test_goto_skip_block() {
    let src = r#"
program T;
label skip;
var
  n: Integer;
begin
  n := 10;
  if n > 5 then goto skip;
  WriteLn('not printed');
  WriteLn('also not printed');
  skip:
  WriteLn('after skip');
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["after skip"]);
}

#[test]
fn test_goto_error_exit() {
    let src = r#"
program T;
label error_exit;
var
  x: Integer;
begin
  x := -1;
  if x < 0 then goto error_exit;
  WriteLn('normal');
  goto error_exit;
  error_exit:
  WriteLn('exit');
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["exit"]);
}

#[test]
fn test_goto_count_down() {
    let src = r#"
program T;
label again;
var
  n: Integer;
begin
  n := 3;
  again:
  WriteLn(n);
  n := n - 1;
  if n > 0 then goto again;
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["3", "2", "1"]);
}

#[test]
fn test_halt_exits_program() {
    let src = r#"
program T;
begin
  WriteLn('before halt');
  Halt;
  WriteLn('after halt');
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["before halt"]);
}

#[test]
fn test_halt_with_code() {
    let src = r#"
program T;
begin
  WriteLn('stopping');
  Halt(0);
  WriteLn('not reached');
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["stopping"]);
}

#[test]
fn test_exit_from_procedure_early() {
    let src = r#"
program T;
procedure Check(n: Integer);
begin
  if n < 0 then begin
    WriteLn('negative');
    Exit;
  end;
  WriteLn('positive');
end;
begin
  Check(5);
  Check(-1);
  Check(0);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["positive", "negative", "positive"]);
}

#[test]
fn test_exit_from_function_early() {
    let src = r#"
program T;
function FindFirst(arr: array of Integer; n, target: Integer): Integer;
var
  i: Integer;
begin
  Result := -1;
  for i := 0 to n - 1 do begin
    if arr[i] = target then begin
      Result := i;
      Exit;
    end;
  end;
end;
var
  data: array[0..4] of Integer;
begin
  data[0] := 3; data[1] := 7; data[2] := 2; data[3] := 7; data[4] := 1;
  WriteLn(FindFirst(data, 5, 7));
  WriteLn(FindFirst(data, 5, 9));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["1", "-1"]);
}

#[test]
fn test_break_in_nested_for() {
    let src = r#"
program T;
var
  i, j: Integer;
begin
  for i := 1 to 3 do begin
    for j := 1 to 3 do begin
      if j = 2 then Break;
      WriteLn(IntToStr(i) + ',' + IntToStr(j));
    end;
  end;
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["1,1", "2,1", "3,1"]);
}

#[test]
fn test_continue_in_for() {
    let src = r#"
program T;
var
  i: Integer;
begin
  for i := 1 to 5 do begin
    if i mod 2 = 0 then Continue;
    WriteLn(i);
  end;
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["1", "3", "5"]);
}

#[test]
fn test_break_in_while() {
    let src = r#"
program T;
var
  n: Integer;
begin
  n := 0;
  while true do begin
    if n >= 3 then Break;
    WriteLn(n);
    n := n + 1;
  end;
  WriteLn('done');
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["0", "1", "2", "done"]);
}

#[test]
fn test_continue_in_while() {
    let src = r#"
program T;
var
  n: Integer;
begin
  n := 0;
  while n < 6 do begin
    n := n + 1;
    if n mod 3 = 0 then Continue;
    WriteLn(n);
  end;
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["1", "2", "4", "5"]);
}

#[test]
fn test_break_in_repeat() {
    let src = r#"
program T;
var
  i: Integer;
begin
  i := 0;
  repeat
    if i = 3 then Break;
    WriteLn(i);
    i := i + 1;
  until false;
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["0", "1", "2"]);
}

#[test]
fn test_continue_skips_rest_of_body() {
    let src = r#"
program T;
var
  i: Integer;
begin
  for i := 1 to 4 do begin
    if i = 3 then Continue;
    Write(IntToStr(i) + ' ');
  end;
  WriteLn('');
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["1 2 4 "]);
}

#[test]
fn test_exit_result_returns_value() {
    let src = r#"
program T;
function Clamp(v, lo, hi: Integer): Integer;
begin
  if v < lo then begin Result := lo; Exit; end;
  if v > hi then begin Result := hi; Exit; end;
  Result := v;
end;
begin
  WriteLn(Clamp(5, 0, 10));
  WriteLn(Clamp(-5, 0, 10));
  WriteLn(Clamp(15, 0, 10));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["5", "0", "10"]);
}

#[test]
fn test_nested_break_only_inner() {
    let src = r#"
program T;
var
  i, j, found: Integer;
begin
  found := -1;
  for i := 1 to 3 do begin
    for j := 1 to 3 do begin
      if i * j = 6 then begin
        found := i * 10 + j;
        Break;
      end;
    end;
    if found > 0 then Break;
  end;
  WriteLn(found);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["23"]);
}

#[test]
fn test_exit_in_nested_if() {
    let src = r#"
program T;
function Classify(n: Integer): string;
begin
  if n < 0 then begin
    Result := 'negative';
    Exit;
  end;
  if n = 0 then begin
    Result := 'zero';
    Exit;
  end;
  if n < 10 then begin
    Result := 'small';
    Exit;
  end;
  Result := 'large';
end;
begin
  WriteLn(Classify(-5));
  WriteLn(Classify(0));
  WriteLn(Classify(7));
  WriteLn(Classify(100));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["negative", "zero", "small", "large"]);
}

#[test]
fn test_goto_two_labels() {
    let src = r#"
program T;
label lblA, lblB;
var
  x: Integer;
begin
  x := 2;
  if x = 1 then goto lblA;
  if x = 2 then goto lblB;
  WriteLn('default');
  goto lblA;
  lblB:
  WriteLn('B');
  goto lblA;
  lblA:
  WriteLn('A');
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["B", "A"]);
}

#[test]
fn test_break_accumulate_before_break() {
    let src = r#"
program T;
var
  i, sum: Integer;
begin
  sum := 0;
  for i := 1 to 100 do begin
    sum := sum + i;
    if sum > 20 then Break;
  end;
  WriteLn(i);
  WriteLn(sum);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["6", "21"]);
}
