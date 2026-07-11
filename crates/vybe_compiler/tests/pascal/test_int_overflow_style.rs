/// Integer div, mod, shl, shr edge cases.
use super::helpers::run_pascal;

#[test]
fn int_div_7_2() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(7 div 2); end."#),
        &["3"]
    );
}

#[test]
fn int_mod_7_2() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(7 mod 2); end."#),
        &["1"]
    );
}

#[test]
fn int_div_10_3() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(10 div 3); end."#),
        &["3"]
    );
}

#[test]
fn int_mod_10_3() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(10 mod 3); end."#),
        &["1"]
    );
}

#[test]
fn int_div_15_4() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(15 div 4); end."#),
        &["3"]
    );
}

#[test]
fn int_mod_15_4() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(15 mod 4); end."#),
        &["3"]
    );
}

#[test]
fn int_div_neg7_2() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(-7 div 2); end."#),
        &["-3"]
    );
}

#[test]
fn int_mod_neg7_2() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(-7 mod 2); end."#),
        &["-1"]
    );
}

#[test]
fn int_div_17_5() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(17 div 5); end."#),
        &["3"]
    );
}

#[test]
fn int_mod_17_5() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(17 mod 5); end."#),
        &["2"]
    );
}

#[test]
fn int_div_100_7() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(100 div 7); end."#),
        &["14"]
    );
}

#[test]
fn int_mod_100_7() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(100 mod 7); end."#),
        &["2"]
    );
}

#[test]
fn int_div_1_1() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(1 div 1); end."#),
        &["1"]
    );
}

#[test]
fn int_mod_1_1() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(1 mod 1); end."#),
        &["0"]
    );
}

#[test]
fn int_div_0_5() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(0 div 5); end."#),
        &["0"]
    );
}

#[test]
fn int_mod_0_5() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(0 mod 5); end."#),
        &["0"]
    );
}

#[test]
fn int_div_9_9() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(9 div 9); end."#),
        &["1"]
    );
}

#[test]
fn int_mod_9_9() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(9 mod 9); end."#),
        &["0"]
    );
}

#[test]
fn int_div_20_6() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(20 div 6); end."#),
        &["3"]
    );
}

#[test]
fn int_mod_20_6() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(20 mod 6); end."#),
        &["2"]
    );
}

#[test]
fn shl_0() {
    assert_eq!(
        run_pascal(r#"program T; var n:Integer; begin n:=1; WriteLn(n shl 0); end."#),
        &["1"]
    );
}

#[test]
fn shl_1() {
    assert_eq!(
        run_pascal(r#"program T; var n:Integer; begin n:=1; WriteLn(n shl 1); end."#),
        &["2"]
    );
}

#[test]
fn shl_2() {
    assert_eq!(
        run_pascal(r#"program T; var n:Integer; begin n:=1; WriteLn(n shl 2); end."#),
        &["4"]
    );
}

#[test]
fn shl_3() {
    assert_eq!(
        run_pascal(r#"program T; var n:Integer; begin n:=1; WriteLn(n shl 3); end."#),
        &["8"]
    );
}

#[test]
fn shl_4() {
    assert_eq!(
        run_pascal(r#"program T; var n:Integer; begin n:=1; WriteLn(n shl 4); end."#),
        &["16"]
    );
}

#[test]
fn shl_5() {
    assert_eq!(
        run_pascal(r#"program T; var n:Integer; begin n:=1; WriteLn(n shl 5); end."#),
        &["32"]
    );
}

#[test]
fn shl_6() {
    assert_eq!(
        run_pascal(r#"program T; var n:Integer; begin n:=1; WriteLn(n shl 6); end."#),
        &["64"]
    );
}

#[test]
fn shl_7() {
    assert_eq!(
        run_pascal(r#"program T; var n:Integer; begin n:=1; WriteLn(n shl 7); end."#),
        &["128"]
    );
}

#[test]
fn shl_8() {
    assert_eq!(
        run_pascal(r#"program T; var n:Integer; begin n:=1; WriteLn(n shl 8); end."#),
        &["256"]
    );
}

#[test]
fn shl_9() {
    assert_eq!(
        run_pascal(r#"program T; var n:Integer; begin n:=1; WriteLn(n shl 9); end."#),
        &["512"]
    );
}

#[test]
fn shl_10() {
    assert_eq!(
        run_pascal(r#"program T; var n:Integer; begin n:=1; WriteLn(n shl 10); end."#),
        &["1024"]
    );
}

#[test]
fn shl_11() {
    assert_eq!(
        run_pascal(r#"program T; var n:Integer; begin n:=1; WriteLn(n shl 11); end."#),
        &["2048"]
    );
}

#[test]
fn shr_1() {
    assert_eq!(
        run_pascal(r#"program T; var n:Integer; begin n:=1024; WriteLn(n shr 1); end."#),
        &["512"]
    );
}

#[test]
fn shr_2() {
    assert_eq!(
        run_pascal(r#"program T; var n:Integer; begin n:=1024; WriteLn(n shr 2); end."#),
        &["256"]
    );
}

#[test]
fn shr_3() {
    assert_eq!(
        run_pascal(r#"program T; var n:Integer; begin n:=1024; WriteLn(n shr 3); end."#),
        &["128"]
    );
}

#[test]
fn shr_4() {
    assert_eq!(
        run_pascal(r#"program T; var n:Integer; begin n:=1024; WriteLn(n shr 4); end."#),
        &["64"]
    );
}

#[test]
fn shr_5() {
    assert_eq!(
        run_pascal(r#"program T; var n:Integer; begin n:=1024; WriteLn(n shr 5); end."#),
        &["32"]
    );
}

#[test]
fn shr_6() {
    assert_eq!(
        run_pascal(r#"program T; var n:Integer; begin n:=1024; WriteLn(n shr 6); end."#),
        &["16"]
    );
}

#[test]
fn shr_7() {
    assert_eq!(
        run_pascal(r#"program T; var n:Integer; begin n:=1024; WriteLn(n shr 7); end."#),
        &["8"]
    );
}

#[test]
fn shr_8() {
    assert_eq!(
        run_pascal(r#"program T; var n:Integer; begin n:=1024; WriteLn(n shr 8); end."#),
        &["4"]
    );
}

#[test]
fn shr_9() {
    assert_eq!(
        run_pascal(r#"program T; var n:Integer; begin n:=1024; WriteLn(n shr 9); end."#),
        &["2"]
    );
}

#[test]
fn shr_10() {
    assert_eq!(
        run_pascal(r#"program T; var n:Integer; begin n:=1024; WriteLn(n shr 10); end."#),
        &["1"]
    );
}
