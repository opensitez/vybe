/// Deep nesting and mutual recursion among nested procedures.
use super::helpers::run_pascal;

#[test]
fn nested_depth_2() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Outer; procedure I1; begin WriteLn(2); end;  begin I1; end; begin Outer; end."#
        ),
        &["2"]
    );
}

#[test]
fn nested_depth_3() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Outer; procedure I1; begin I2; end; procedure I2; begin WriteLn(3); end;  begin I1; end; begin Outer; end."#
        ),
        &["3"]
    );
}

#[test]
fn nested_depth_4() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Outer; procedure I1; begin I2; end; procedure I2; begin I3; end; procedure I3; begin WriteLn(4); end;  begin I1; end; begin Outer; end."#
        ),
        &["4"]
    );
}

#[test]
fn nested_depth_5() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Outer; procedure I1; begin I2; end; procedure I2; begin I3; end; procedure I3; begin I4; end; procedure I4; begin WriteLn(5); end;  begin I1; end; begin Outer; end."#
        ),
        &["5"]
    );
}

#[test]
fn nested_depth_6() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Outer; procedure I1; begin I2; end; procedure I2; begin I3; end; procedure I3; begin I4; end; procedure I4; begin I5; end; procedure I5; begin WriteLn(6); end;  begin I1; end; begin Outer; end."#
        ),
        &["6"]
    );
}

#[test]
fn nested_depth_7() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Outer; procedure I1; begin I2; end; procedure I2; begin I3; end; procedure I3; begin I4; end; procedure I4; begin I5; end; procedure I5; begin I6; end; procedure I6; begin WriteLn(7); end;  begin I1; end; begin Outer; end."#
        ),
        &["7"]
    );
}

#[test]
fn nested_depth_8() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Outer; procedure I1; begin I2; end; procedure I2; begin I3; end; procedure I3; begin I4; end; procedure I4; begin I5; end; procedure I5; begin I6; end; procedure I6; begin I7; end; procedure I7; begin WriteLn(8); end;  begin I1; end; begin Outer; end."#
        ),
        &["8"]
    );
}

#[test]
fn nested_depth_9() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Outer; procedure I1; begin I2; end; procedure I2; begin I3; end; procedure I3; begin I4; end; procedure I4; begin I5; end; procedure I5; begin I6; end; procedure I6; begin I7; end; procedure I7; begin I8; end; procedure I8; begin WriteLn(9); end;  begin I1; end; begin Outer; end."#
        ),
        &["9"]
    );
}

#[test]
fn nested_depth_10() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Outer; procedure I1; begin I2; end; procedure I2; begin I3; end; procedure I3; begin I4; end; procedure I4; begin I5; end; procedure I5; begin I6; end; procedure I6; begin I7; end; procedure I7; begin I8; end; procedure I8; begin I9; end; procedure I9; begin WriteLn(10); end;  begin I1; end; begin Outer; end."#
        ),
        &["10"]
    );
}

#[test]
fn nested_depth_11() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Outer; procedure I1; begin I2; end; procedure I2; begin I3; end; procedure I3; begin I4; end; procedure I4; begin I5; end; procedure I5; begin I6; end; procedure I6; begin I7; end; procedure I7; begin I8; end; procedure I8; begin I9; end; procedure I9; begin I10; end; procedure I10; begin WriteLn(11); end;  begin I1; end; begin Outer; end."#
        ),
        &["11"]
    );
}

#[test]
fn nested_depth_12() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Outer; procedure I1; begin I2; end; procedure I2; begin I3; end; procedure I3; begin I4; end; procedure I4; begin I5; end; procedure I5; begin I6; end; procedure I6; begin I7; end; procedure I7; begin I8; end; procedure I8; begin I9; end; procedure I9; begin I10; end; procedure I10; begin I11; end; procedure I11; begin WriteLn(12); end;  begin I1; end; begin Outer; end."#
        ),
        &["12"]
    );
}

#[test]
fn nested_depth_13() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Outer; procedure I1; begin I2; end; procedure I2; begin I3; end; procedure I3; begin I4; end; procedure I4; begin I5; end; procedure I5; begin I6; end; procedure I6; begin I7; end; procedure I7; begin I8; end; procedure I8; begin I9; end; procedure I9; begin I10; end; procedure I10; begin I11; end; procedure I11; begin I12; end; procedure I12; begin WriteLn(13); end;  begin I1; end; begin Outer; end."#
        ),
        &["13"]
    );
}

#[test]
fn nested_depth_14() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Outer; procedure I1; begin I2; end; procedure I2; begin I3; end; procedure I3; begin I4; end; procedure I4; begin I5; end; procedure I5; begin I6; end; procedure I6; begin I7; end; procedure I7; begin I8; end; procedure I8; begin I9; end; procedure I9; begin I10; end; procedure I10; begin I11; end; procedure I11; begin I12; end; procedure I12; begin I13; end; procedure I13; begin WriteLn(14); end;  begin I1; end; begin Outer; end."#
        ),
        &["14"]
    );
}

#[test]
fn nested_depth_15() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Outer; procedure I1; begin I2; end; procedure I2; begin I3; end; procedure I3; begin I4; end; procedure I4; begin I5; end; procedure I5; begin I6; end; procedure I6; begin I7; end; procedure I7; begin I8; end; procedure I8; begin I9; end; procedure I9; begin I10; end; procedure I10; begin I11; end; procedure I11; begin I12; end; procedure I12; begin I13; end; procedure I13; begin I14; end; procedure I14; begin WriteLn(15); end;  begin I1; end; begin Outer; end."#
        ),
        &["15"]
    );
}

#[test]
fn mutual_recursion_countdown_1() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure B(n:Integer); forward; procedure A(n:Integer); begin if n>0 then begin WriteLn(n); B(n-1); end; end; procedure B(n:Integer); begin if n>0 then A(n-1); end; begin A(1); end."#
        ),
        &["1"]
    );
}

#[test]
fn mutual_recursion_countdown_2() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure B(n:Integer); forward; procedure A(n:Integer); begin if n>0 then begin WriteLn(n); B(n-1); end; end; procedure B(n:Integer); begin if n>0 then A(n-1); end; begin A(2); end."#
        ),
        &["2"]
    );
}

#[test]
fn mutual_recursion_countdown_3() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure B(n:Integer); forward; procedure A(n:Integer); begin if n>0 then begin WriteLn(n); B(n-1); end; end; procedure B(n:Integer); begin if n>0 then A(n-1); end; begin A(3); end."#
        ),
        &["3", "1"]
    );
}

#[test]
fn mutual_recursion_countdown_4() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure B(n:Integer); forward; procedure A(n:Integer); begin if n>0 then begin WriteLn(n); B(n-1); end; end; procedure B(n:Integer); begin if n>0 then A(n-1); end; begin A(4); end."#
        ),
        &["4", "2"]
    );
}

#[test]
fn mutual_recursion_countdown_5() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure B(n:Integer); forward; procedure A(n:Integer); begin if n>0 then begin WriteLn(n); B(n-1); end; end; procedure B(n:Integer); begin if n>0 then A(n-1); end; begin A(5); end."#
        ),
        &["5", "3", "1"]
    );
}

#[test]
fn mutual_recursion_countdown_6() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure B(n:Integer); forward; procedure A(n:Integer); begin if n>0 then begin WriteLn(n); B(n-1); end; end; procedure B(n:Integer); begin if n>0 then A(n-1); end; begin A(6); end."#
        ),
        &["6", "4", "2"]
    );
}

#[test]
fn mutual_recursion_countdown_7() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure B(n:Integer); forward; procedure A(n:Integer); begin if n>0 then begin WriteLn(n); B(n-1); end; end; procedure B(n:Integer); begin if n>0 then A(n-1); end; begin A(7); end."#
        ),
        &["7", "5", "3", "1"]
    );
}

#[test]
fn mutual_recursion_countdown_8() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure B(n:Integer); forward; procedure A(n:Integer); begin if n>0 then begin WriteLn(n); B(n-1); end; end; procedure B(n:Integer); begin if n>0 then A(n-1); end; begin A(8); end."#
        ),
        &["8", "6", "4", "2"]
    );
}

#[test]
fn mutual_recursion_countdown_9() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure B(n:Integer); forward; procedure A(n:Integer); begin if n>0 then begin WriteLn(n); B(n-1); end; end; procedure B(n:Integer); begin if n>0 then A(n-1); end; begin A(9); end."#
        ),
        &["9", "7", "5", "3", "1"]
    );
}

#[test]
fn mutual_recursion_countdown_10() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure B(n:Integer); forward; procedure A(n:Integer); begin if n>0 then begin WriteLn(n); B(n-1); end; end; procedure B(n:Integer); begin if n>0 then A(n-1); end; begin A(10); end."#
        ),
        &["10", "8", "6", "4", "2"]
    );
}

#[test]
fn mutual_recursion_countdown_11() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure B(n:Integer); forward; procedure A(n:Integer); begin if n>0 then begin WriteLn(n); B(n-1); end; end; procedure B(n:Integer); begin if n>0 then A(n-1); end; begin A(11); end."#
        ),
        &["11", "9", "7", "5", "3", "1"]
    );
}

#[test]
fn mutual_recursion_countdown_12() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure B(n:Integer); forward; procedure A(n:Integer); begin if n>0 then begin WriteLn(n); B(n-1); end; end; procedure B(n:Integer); begin if n>0 then A(n-1); end; begin A(12); end."#
        ),
        &["12", "10", "8", "6", "4", "2"]
    );
}

#[test]
fn mutual_recursion_countdown_13() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure B(n:Integer); forward; procedure A(n:Integer); begin if n>0 then begin WriteLn(n); B(n-1); end; end; procedure B(n:Integer); begin if n>0 then A(n-1); end; begin A(13); end."#
        ),
        &["13", "11", "9", "7", "5", "3", "1"]
    );
}

#[test]
fn mutual_recursion_countdown_14() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure B(n:Integer); forward; procedure A(n:Integer); begin if n>0 then begin WriteLn(n); B(n-1); end; end; procedure B(n:Integer); begin if n>0 then A(n-1); end; begin A(14); end."#
        ),
        &["14", "12", "10", "8", "6", "4", "2"]
    );
}

#[test]
fn nested_func_returns_to_outer() {
    assert_eq!(
        run_pascal(
            r#"program T; function Outer:Integer; function Inner:Integer; begin Result:=42; end; begin Result:=Inner; end; begin WriteLn(Outer); end."#
        ),
        &["42"]
    );
}

#[test]
fn nested_proc_modifies_outer_local() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Outer; var x:Integer; procedure Inner; begin x:=x+5; end; begin x:=10; Inner; WriteLn(x); end; begin Outer; end."#
        ),
        &["15"]
    );
}

#[test]
fn nested_two_procs_same_level() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Outer; procedure A; begin WriteLn(1); end; procedure B; begin WriteLn(2); end; begin A; B; end; begin Outer; end."#
        ),
        &["1", "2"]
    );
}

#[test]
fn nested_proc_reads_outer_const() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Outer; const K=11; procedure Inner; begin WriteLn(K); end; begin Inner; end; begin Outer; end."#
        ),
        &["11"]
    );
}

#[test]
fn nested_three_level_prints_chain() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure A; procedure B; procedure C; begin WriteLn('c'); end; begin C; WriteLn('b'); end; begin B; WriteLn('a'); end; begin A; end."#
        ),
        &["c", "b", "a"]
    );
}

#[test]
fn nested_func_uses_outer_param() {
    assert_eq!(
        run_pascal(
            r#"program T; function Outer(n:Integer):Integer; function Inner:Integer; begin Result:=n*2; end; begin Result:=Inner; end; begin WriteLn(Outer(6)); end."#
        ),
        &["12"]
    );
}

#[test]
fn nested_proc_shadows_outer_var() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Outer; var x:Integer; procedure Inner; var x:Integer; begin x:=99; WriteLn(x); end; begin x:=1; Inner; WriteLn(x); end; begin Outer; end."#
        ),
        &["99", "1"]
    );
}

#[test]
fn nested_four_funcs_sum() {
    assert_eq!(
        run_pascal(
            r#"program T; function Outer:Integer; function A:Integer; begin Result:=1; end; function B:Integer; begin Result:=A+2; end; function C:Integer; begin Result:=B+3; end; begin Result:=C; end; begin WriteLn(Outer); end."#
        ),
        &["6"]
    );
}

#[test]
fn nested_proc_called_twice() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Outer; procedure Tick; begin WriteLn('t'); end; begin Tick; Tick; end; begin Outer; end."#
        ),
        &["t", "t"]
    );
}

#[test]
fn nested_proc_accumulates_sum() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Outer; var total:Integer; procedure Add(n:Integer); begin total:=total+n; end; begin total:=0; Add(3); Add(4); WriteLn(total); end; begin Outer; end."#
        ),
        &["7"]
    );
}

#[test]
fn nested_proc_in_function_body() {
    assert_eq!(
        run_pascal(
            r#"program T; function Compute:Integer; procedure Step; begin WriteLn(1); end; begin Step; Result:=7; end; begin WriteLn(Compute); end."#
        ),
        &["1", "7"]
    );
}

#[test]
fn nested_five_level_accumulator() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure L1; procedure L2; procedure L3; procedure L4; procedure L5; begin WriteLn(5); end; begin L5; end; begin L4; end; begin L3; end; begin L2; end; begin L1; end."#
        ),
        &["5"]
    );
}

#[test]
fn nested_mutual_local_forward() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Outer; procedure B; forward; procedure A; begin WriteLn('a'); B; end; procedure B; begin WriteLn('b'); end; begin A; end; begin Outer; end."#
        ),
        &["a", "b"]
    );
}

#[test]
fn nested_deep_six_with_output() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure L1; procedure L2; procedure L3; procedure L4; procedure L5; procedure L6; begin WriteLn(6); end; begin L6; end; begin L5; end; begin L4; end; begin L3; end; begin L2; end; begin L1; end."#
        ),
        &["6"]
    );
}

