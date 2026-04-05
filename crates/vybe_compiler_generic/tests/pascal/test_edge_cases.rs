
use super::helpers::run;

#[test] fn edge_empty_program()    { assert_eq!(run("program T; begin end."), &[] as &[&str]); }
#[test] fn edge_empty_proc()       { assert_eq!(run("program T; procedure P; begin end; begin P; end."), &[] as &[&str]); }
#[test] fn edge_many_locals() {
    assert_eq!(run("program T; var a,b,c,d,e: Integer; begin a:=1; b:=2; c:=3; d:=4; e:=5; WriteLn(a+b+c+d+e); end."), &["15"]);
}
#[test] fn edge_deeply_nested_calls() {
    assert_eq!(run(r#"program T;
function F(x: Integer): Integer; begin Result := x + 1; end;
begin WriteLn(F(F(F(F(F(0)))))); end."#), &["5"]);
}
#[test] fn edge_zero_iterations() {
    assert_eq!(run("program T; var i: Integer; begin for i := 5 to 3 do WriteLn('x'); WriteLn('done'); end."), &["done"]);
}
#[test] fn edge_single_char_string() {
    assert_eq!(run("program T; begin WriteLn(Length('x')); end."), &["1"]);
}
#[test] fn edge_assigned_nil() {
    assert_eq!(run("program T; begin if not Assigned(nil) then WriteLn('y'); end."), &["y"]);
}

// Scoping
#[test] fn scope_local_shadows_global() {
    assert_eq!(run(r#"program T; var x: Integer;
procedure Test; var x: Integer; begin x := 99; WriteLn(x); end;
begin x := 1; Test; WriteLn(x); end."#), &["99", "1"]);
}
#[test] fn scope_nested_blocks() {
    assert_eq!(run("program T; var x: Integer; begin x := 1; begin x := 2; WriteLn(x); end; WriteLn(x); end."), &["2", "2"]);
}

// Multiple functions interacting
#[test] fn funcs_calling_each_other() {
    assert_eq!(run(r#"program T;
function A(x: Integer): Integer; begin Result := x * 2; end;
function B(x: Integer): Integer; begin Result := A(x) + 1; end;
begin WriteLn(B(5)); end."#), &["11"]);
}

// Function result used in comparison
#[test] fn func_result_in_if() {
    assert_eq!(run(r#"program T;
function IsPositive(x: Integer): Boolean; begin Result := x > 0; end;
begin if IsPositive(5) then WriteLn('pos') else WriteLn('neg'); end."#), &["pos"]);
}

// Long string operations
#[test] fn long_string_build() {
    assert_eq!(run(r#"program T; var s: String; var i: Integer;
begin s := ''; for i := 1 to 5 do s := s + IntToStr(i); WriteLn(s); end."#), &["12345"]);
}

// Multiple WriteLn with expressions
#[test] fn mixed_output() {
    assert_eq!(run(r#"program T; begin
      WriteLn(1 + 2);
      WriteLn('hello');
      WriteLn(true);
      WriteLn(3.14);
    end."#), &["3", "hello", "true", "3.14"]);
}

// Nested loops with accumulator
#[test] fn nested_loop_sum() {
    assert_eq!(run(r#"program T; var i, j, s: Integer;
begin s := 0;
  for i := 1 to 3 do for j := 1 to 3 do s := s + i * j;
  WriteLn(s); end."#), &["36"]);
}

// While loop with complex condition
#[test] fn while_complex_cond() {
    assert_eq!(run(r#"program T; var x: Integer;
begin x := 100;
  while (x > 1) and (x > 50) do x := x - 10;
  WriteLn(x); end."#), &["50"]);
}
