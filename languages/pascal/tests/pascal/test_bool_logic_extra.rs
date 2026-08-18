/// and/or/not/xor and short-circuit boolean patterns.
use super::helpers::run_pascal;

#[test]
fn and_both_true() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(true and true); end."#),
        &["TRUE"]
    );
}

#[test]
fn and_left_false() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(false and true); end."#),
        &["FALSE"]
    );
}

#[test]
fn and_right_false() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(true and false); end."#),
        &["FALSE"]
    );
}

#[test]
fn or_left_true() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(true or false); end."#),
        &["TRUE"]
    );
}

#[test]
fn or_both_false() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(false or false); end."#),
        &["FALSE"]
    );
}

#[test]
fn or_right_true() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(false or true); end."#),
        &["TRUE"]
    );
}

#[test]
fn not_true() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(not true); end."#),
        &["FALSE"]
    );
}

#[test]
fn not_false() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(not false); end."#),
        &["TRUE"]
    );
}

#[test]
fn xor_diff() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(true xor false); end."#),
        &["true"]
    );
}

#[test]
fn xor_same() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(true xor true); end."#),
        &["false"]
    );
}

#[test]
fn xor_false_true() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(false xor true); end."#),
        &["true"]
    );
}

#[test]
fn and_short_circuit() {
    assert_eq!(
        run_pascal(
            r#"program T; var n:Integer; begin n:=0; if false and (n=1) then WriteLn('y') else WriteLn('n'); end."#
        ),
        &["n"]
    );
}

#[test]
fn or_short_circuit() {
    assert_eq!(
        run_pascal(
            r#"program T; begin if true or (1>2) then WriteLn('y') else WriteLn('n'); end."#
        ),
        &["y"]
    );
}

#[test]
fn and_chain() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn((1<2) and (3<4) and (5<6)); end."#),
        &["TRUE"]
    );
}

#[test]
fn or_chain() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn((1>2) or (3>4) or (5<6)); end."#),
        &["TRUE"]
    );
}

#[test]
fn not_and_demorgan() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Boolean; begin a:=true; b:=false; WriteLn(not (a and b)); end."#
        ),
        &["TRUE"]
    );
}

#[test]
fn not_or_demorgan() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Boolean; begin a:=false; b:=false; WriteLn(not (a or b)); end."#
        ),
        &["TRUE"]
    );
}

#[test]
fn if_and_condition() {
    assert_eq!(
        run_pascal(r#"program T; begin if (2>1) and (3>2) then WriteLn('ok'); end."#),
        &["ok"]
    );
}

#[test]
fn if_or_condition() {
    assert_eq!(
        run_pascal(r#"program T; begin if (2>3) or (4>3) then WriteLn('ok'); end."#),
        &["ok"]
    );
}

#[test]
fn if_xor_condition() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Boolean; begin a:=true; b:=false; if a xor b then WriteLn('xor'); end."#
        ),
        &["xor"]
    );
}

#[test]
fn while_and_guard() {
    assert_eq!(
        run_pascal(
            r#"program T; var i:Integer; begin i:=0; while (i<3) and (i>=0) do begin WriteLn(i); Inc(i); end; end."#
        ),
        &["0", "1", "2"]
    );
}

#[test]
fn repeat_until_or() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Integer; begin a:=0; b:=10; repeat Inc(a); until (a>=3) or (b<0); WriteLn(a); end."#
        ),
        &["3"]
    );
}

#[test]
fn repeat_until_and() {
    assert_eq!(
        run_pascal(
            r#"program T; var x,y:Integer; begin x:=0; y:=0; repeat Inc(x); Inc(y); until (x>=2) and (y>=2); WriteLn(x); end."#
        ),
        &["2"]
    );
}

#[test]
fn boolean_var_assign() {
    assert_eq!(
        run_pascal(r#"program T; var f:Boolean; begin f:=3>2; WriteLn(f); end."#),
        &["TRUE"]
    );
}

#[test]
fn nested_not() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(not not true); end."#),
        &["TRUE"]
    );
}

#[test]
fn xor_in_loop_parity() {
    assert_eq!(
        run_pascal(
            r#"program T; var i,p:Integer; begin p:=0; for i:=1 to 3 do p:=p xor 1; WriteLn(p); end."#
        ),
        &["1"]
    );
}

#[test]
fn and_with_compare() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn((5=5) and (6>5)); end."#),
        &["TRUE"]
    );
}

#[test]
fn or_with_compare() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn((5=6) or (7=7)); end."#),
        &["TRUE"]
    );
}

#[test]
fn if_not_or() {
    assert_eq!(
        run_pascal(r#"program T; begin if not (false or false) then WriteLn('ok'); end."#),
        &["ok"]
    );
}

#[test]
fn complex_paren() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn((true or false) and (not false)); end."#),
        &["TRUE"]
    );
}

#[test]
fn xor_three_terms() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn((true xor false) xor false); end."#),
        &["true"]
    );
}

#[test]
fn bool_from_equal() {
    assert_eq!(
        run_pascal(r#"program T; var a,b:Integer; begin a:=4; b:=4; WriteLn(a=b); end."#),
        &["TRUE"]
    );
}

#[test]
fn bool_from_not_equal() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(3<>4); end."#),
        &["TRUE"]
    );
}

#[test]
fn if_and_else_or() {
    assert_eq!(
        run_pascal(
            r#"program T; begin if (1<0) and (2<3) then WriteLn('a') else if (3<4) or (5<6) then WriteLn('b'); end."#
        ),
        &["b"]
    );
}

#[test]
fn while_not_condition() {
    assert_eq!(
        run_pascal(
            r#"program T; var i:Integer; begin i:=0; while not (i>2) do begin WriteLn(i); Inc(i); end; end."#
        ),
        &["0", "1", "2"]
    );
}

#[test]
fn for_with_bool_break() {
    assert_eq!(
        run_pascal(
            r#"program T; var i:Integer; done:Boolean; begin done:=false; for i:=1 to 5 do if not done then begin if i=3 then done:=true else WriteLn(i); end; end."#
        ),
        &["1", "2"]
    );
}

#[test]
fn and_false_stops() {
    assert_eq!(
        run_pascal(
            r#"program T; begin if (2>1) and (9<1) then WriteLn('t') else WriteLn('f'); end."#
        ),
        &["f"]
    );
}

#[test]
fn or_true_stops() {
    assert_eq!(
        run_pascal(
            r#"program T; begin if (9<1) or (2>1) then WriteLn('t') else WriteLn('f'); end."#
        ),
        &["t"]
    );
}

#[test]
fn not_equal_in_if() {
    assert_eq!(
        run_pascal(r#"program T; var s:string; begin s:='a'; if s<>'b' then WriteLn('ne'); end."#),
        &["ne"]
    );
}

#[test]
fn xor_zero_result() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(false xor false); end."#),
        &["false"]
    );
}

#[test]
fn boolean_implies_style() {
    assert_eq!(
        run_pascal(
            r#"program T; begin if (not true) or (false) then WriteLn('imp') else WriteLn('fail'); end."#
        ),
        &["fail"]
    );
}

#[test]
fn triple_and() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(true and true and true); end."#),
        &["TRUE"]
    );
}
