/// Goto and label control flow patterns.
use super::helpers::run_pascal;

#[test]
fn goto_count_up_to_1() {
    assert_eq!(
        run_pascal(
            r#"program T; label top; var i:Integer; begin i:=0; top: Inc(i); if i<1 then goto top; WriteLn(i); end."#
        ),
        &["1"]
    );
}

#[test]
fn goto_count_up_to_2() {
    assert_eq!(
        run_pascal(
            r#"program T; label top; var i:Integer; begin i:=0; top: Inc(i); if i<2 then goto top; WriteLn(i); end."#
        ),
        &["2"]
    );
}

#[test]
fn goto_count_up_to_3() {
    assert_eq!(
        run_pascal(
            r#"program T; label top; var i:Integer; begin i:=0; top: Inc(i); if i<3 then goto top; WriteLn(i); end."#
        ),
        &["3"]
    );
}

#[test]
fn goto_count_up_to_4() {
    assert_eq!(
        run_pascal(
            r#"program T; label top; var i:Integer; begin i:=0; top: Inc(i); if i<4 then goto top; WriteLn(i); end."#
        ),
        &["4"]
    );
}

#[test]
fn goto_count_up_to_5() {
    assert_eq!(
        run_pascal(
            r#"program T; label top; var i:Integer; begin i:=0; top: Inc(i); if i<5 then goto top; WriteLn(i); end."#
        ),
        &["5"]
    );
}

#[test]
fn goto_count_up_to_6() {
    assert_eq!(
        run_pascal(
            r#"program T; label top; var i:Integer; begin i:=0; top: Inc(i); if i<6 then goto top; WriteLn(i); end."#
        ),
        &["6"]
    );
}

#[test]
fn goto_count_up_to_7() {
    assert_eq!(
        run_pascal(
            r#"program T; label top; var i:Integer; begin i:=0; top: Inc(i); if i<7 then goto top; WriteLn(i); end."#
        ),
        &["7"]
    );
}

#[test]
fn goto_count_up_to_8() {
    assert_eq!(
        run_pascal(
            r#"program T; label top; var i:Integer; begin i:=0; top: Inc(i); if i<8 then goto top; WriteLn(i); end."#
        ),
        &["8"]
    );
}

#[test]
fn goto_count_up_to_9() {
    assert_eq!(
        run_pascal(
            r#"program T; label top; var i:Integer; begin i:=0; top: Inc(i); if i<9 then goto top; WriteLn(i); end."#
        ),
        &["9"]
    );
}

#[test]
fn goto_count_up_to_10() {
    assert_eq!(
        run_pascal(
            r#"program T; label top; var i:Integer; begin i:=0; top: Inc(i); if i<10 then goto top; WriteLn(i); end."#
        ),
        &["10"]
    );
}

#[test]
fn goto_count_up_to_11() {
    assert_eq!(
        run_pascal(
            r#"program T; label top; var i:Integer; begin i:=0; top: Inc(i); if i<11 then goto top; WriteLn(i); end."#
        ),
        &["11"]
    );
}

#[test]
fn goto_count_up_to_12() {
    assert_eq!(
        run_pascal(
            r#"program T; label top; var i:Integer; begin i:=0; top: Inc(i); if i<12 then goto top; WriteLn(i); end."#
        ),
        &["12"]
    );
}

#[test]
fn goto_count_up_to_13() {
    assert_eq!(
        run_pascal(
            r#"program T; label top; var i:Integer; begin i:=0; top: Inc(i); if i<13 then goto top; WriteLn(i); end."#
        ),
        &["13"]
    );
}

#[test]
fn goto_count_up_to_14() {
    assert_eq!(
        run_pascal(
            r#"program T; label top; var i:Integer; begin i:=0; top: Inc(i); if i<14 then goto top; WriteLn(i); end."#
        ),
        &["14"]
    );
}

#[test]
fn goto_if_branch_1() {
    assert_eq!(
        run_pascal(
            r#"program T; label L; var x:Integer; begin x:=1; if x mod 2=0 then goto L; WriteLn(1); goto endL; L: WriteLn(2); endL: end."#
        ),
        &["1"]
    );
}

#[test]
fn goto_if_branch_2() {
    assert_eq!(
        run_pascal(
            r#"program T; label L; var x:Integer; begin x:=2; if x mod 2=0 then goto L; WriteLn(1); goto endL; L: WriteLn(2); endL: end."#
        ),
        &["2"]
    );
}

#[test]
fn goto_if_branch_3() {
    assert_eq!(
        run_pascal(
            r#"program T; label L; var x:Integer; begin x:=3; if x mod 2=0 then goto L; WriteLn(1); goto endL; L: WriteLn(2); endL: end."#
        ),
        &["1"]
    );
}

#[test]
fn goto_if_branch_4() {
    assert_eq!(
        run_pascal(
            r#"program T; label L; var x:Integer; begin x:=4; if x mod 2=0 then goto L; WriteLn(1); goto endL; L: WriteLn(2); endL: end."#
        ),
        &["2"]
    );
}

#[test]
fn goto_if_branch_5() {
    assert_eq!(
        run_pascal(
            r#"program T; label L; var x:Integer; begin x:=5; if x mod 2=0 then goto L; WriteLn(1); goto endL; L: WriteLn(2); endL: end."#
        ),
        &["1"]
    );
}

#[test]
fn goto_if_branch_6() {
    assert_eq!(
        run_pascal(
            r#"program T; label L; var x:Integer; begin x:=6; if x mod 2=0 then goto L; WriteLn(1); goto endL; L: WriteLn(2); endL: end."#
        ),
        &["2"]
    );
}

#[test]
fn goto_if_branch_7() {
    assert_eq!(
        run_pascal(
            r#"program T; label L; var x:Integer; begin x:=7; if x mod 2=0 then goto L; WriteLn(1); goto endL; L: WriteLn(2); endL: end."#
        ),
        &["1"]
    );
}

#[test]
fn goto_if_branch_8() {
    assert_eq!(
        run_pascal(
            r#"program T; label L; var x:Integer; begin x:=8; if x mod 2=0 then goto L; WriteLn(1); goto endL; L: WriteLn(2); endL: end."#
        ),
        &["2"]
    );
}

#[test]
fn goto_if_branch_9() {
    assert_eq!(
        run_pascal(
            r#"program T; label L; var x:Integer; begin x:=9; if x mod 2=0 then goto L; WriteLn(1); goto endL; L: WriteLn(2); endL: end."#
        ),
        &["1"]
    );
}

#[test]
fn goto_if_branch_10() {
    assert_eq!(
        run_pascal(
            r#"program T; label L; var x:Integer; begin x:=10; if x mod 2=0 then goto L; WriteLn(1); goto endL; L: WriteLn(2); endL: end."#
        ),
        &["2"]
    );
}

#[test]
fn goto_if_branch_11() {
    assert_eq!(
        run_pascal(
            r#"program T; label L; var x:Integer; begin x:=11; if x mod 2=0 then goto L; WriteLn(1); goto endL; L: WriteLn(2); endL: end."#
        ),
        &["1"]
    );
}

#[test]
fn goto_if_branch_12() {
    assert_eq!(
        run_pascal(
            r#"program T; label L; var x:Integer; begin x:=12; if x mod 2=0 then goto L; WriteLn(1); goto endL; L: WriteLn(2); endL: end."#
        ),
        &["2"]
    );
}

#[test]
fn goto_if_branch_13() {
    assert_eq!(
        run_pascal(
            r#"program T; label L; var x:Integer; begin x:=13; if x mod 2=0 then goto L; WriteLn(1); goto endL; L: WriteLn(2); endL: end."#
        ),
        &["1"]
    );
}

#[test]
fn goto_if_branch_14() {
    assert_eq!(
        run_pascal(
            r#"program T; label L; var x:Integer; begin x:=14; if x mod 2=0 then goto L; WriteLn(1); goto endL; L: WriteLn(2); endL: end."#
        ),
        &["2"]
    );
}

#[test]
fn goto_repeat_print_10() {
    assert_eq!(
        run_pascal(r#"program T; label top; var n:Integer; begin n:=10; top: WriteLn(n); Dec(n); if n>0 then goto top; end."#),
        &["10", "9", "8", "7", "6", "5", "4", "3", "2", "1"]
    );
}

#[test]
fn goto_repeat_print_11() {
    assert_eq!(
        run_pascal(r#"program T; label top; var n:Integer; begin n:=11; top: WriteLn(n); Dec(n); if n>0 then goto top; end."#),
        &["11", "10", "9", "8", "7", "6", "5", "4", "3", "2", "1"]
    );
}

#[test]
fn goto_repeat_print_12() {
    assert_eq!(
        run_pascal(r#"program T; label top; var n:Integer; begin n:=12; top: WriteLn(n); Dec(n); if n>0 then goto top; end."#),
        &["12", "11", "10", "9", "8", "7", "6", "5", "4", "3", "2", "1"]
    );
}

#[test]
fn goto_repeat_print_13() {
    assert_eq!(
        run_pascal(r#"program T; label top; var n:Integer; begin n:=13; top: WriteLn(n); Dec(n); if n>0 then goto top; end."#),
        &["13", "12", "11", "10", "9", "8", "7", "6", "5", "4", "3", "2", "1"]
    );
}

#[test]
fn goto_repeat_print_14() {
    assert_eq!(
        run_pascal(r#"program T; label top; var n:Integer; begin n:=14; top: WriteLn(n); Dec(n); if n>0 then goto top; end."#),
        &["14", "13", "12", "11", "10", "9", "8", "7", "6", "5", "4", "3", "2", "1"]
    );
}

#[test]
fn goto_repeat_print_15() {
    assert_eq!(
        run_pascal(r#"program T; label top; var n:Integer; begin n:=15; top: WriteLn(n); Dec(n); if n>0 then goto top; end."#),
        &["15", "14", "13", "12", "11", "10", "9", "8", "7", "6", "5", "4", "3", "2", "1"]
    );
}

#[test]
fn goto_repeat_print_16() {
    assert_eq!(
        run_pascal(r#"program T; label top; var n:Integer; begin n:=16; top: WriteLn(n); Dec(n); if n>0 then goto top; end."#),
        &["16", "15", "14", "13", "12", "11", "10", "9", "8", "7", "6", "5", "4", "3", "2", "1"]
    );
}

#[test]
fn goto_repeat_print_17() {
    assert_eq!(
        run_pascal(r#"program T; label top; var n:Integer; begin n:=17; top: WriteLn(n); Dec(n); if n>0 then goto top; end."#),
        &["17", "16", "15", "14", "13", "12", "11", "10", "9", "8", "7", "6", "5", "4", "3", "2", "1"]
    );
}

#[test]
fn goto_repeat_print_18() {
    assert_eq!(
        run_pascal(r#"program T; label top; var n:Integer; begin n:=18; top: WriteLn(n); Dec(n); if n>0 then goto top; end."#),
        &["18", "17", "16", "15", "14", "13", "12", "11", "10", "9", "8", "7", "6", "5", "4", "3", "2", "1"]
    );
}

#[test]
fn goto_repeat_print_19() {
    assert_eq!(
        run_pascal(r#"program T; label top; var n:Integer; begin n:=19; top: WriteLn(n); Dec(n); if n>0 then goto top; end."#),
        &["19", "18", "17", "16", "15", "14", "13", "12", "11", "10", "9", "8", "7", "6", "5", "4", "3", "2", "1"]
    );
}

#[test]
fn goto_repeat_print_20() {
    assert_eq!(
        run_pascal(r#"program T; label top; var n:Integer; begin n:=20; top: WriteLn(n); Dec(n); if n>0 then goto top; end."#),
        &["20", "19", "18", "17", "16", "15", "14", "13", "12", "11", "10", "9", "8", "7", "6", "5", "4", "3", "2", "1"]
    );
}

#[test]
fn goto_repeat_print_21() {
    assert_eq!(
        run_pascal(r#"program T; label top; var n:Integer; begin n:=21; top: WriteLn(n); Dec(n); if n>0 then goto top; end."#),
        &["21", "20", "19", "18", "17", "16", "15", "14", "13", "12", "11", "10", "9", "8", "7", "6", "5", "4", "3", "2", "1"]
    );
}

#[test]
fn goto_repeat_print_22() {
    assert_eq!(
        run_pascal(r#"program T; label top; var n:Integer; begin n:=22; top: WriteLn(n); Dec(n); if n>0 then goto top; end."#),
        &["22", "21", "20", "19", "18", "17", "16", "15", "14", "13", "12", "11", "10", "9", "8", "7", "6", "5", "4", "3", "2", "1"]
    );
}

#[test]
fn goto_repeat_print_23() {
    assert_eq!(
        run_pascal(r#"program T; label top; var n:Integer; begin n:=23; top: WriteLn(n); Dec(n); if n>0 then goto top; end."#),
        &["23", "22", "21", "20", "19", "18", "17", "16", "15", "14", "13", "12", "11", "10", "9", "8", "7", "6", "5", "4", "3", "2", "1"]
    );
}

