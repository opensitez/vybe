/// Double, Extended, and Currency arithmetic operations.
use super::helpers::run_pascal;

#[test]
fn double_add_1() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Double; begin a:=1.5; b:=1.25; WriteLn(Trunc(a+b)); end."#
        ),
        &["2"]
    );
}

#[test]
fn double_add_2() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Double; begin a:=2.5; b:=2.25; WriteLn(Trunc(a+b)); end."#
        ),
        &["4"]
    );
}

#[test]
fn double_add_3() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Double; begin a:=3.5; b:=3.25; WriteLn(Trunc(a+b)); end."#
        ),
        &["6"]
    );
}

#[test]
fn double_add_4() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Double; begin a:=4.5; b:=4.25; WriteLn(Trunc(a+b)); end."#
        ),
        &["8"]
    );
}

#[test]
fn double_add_5() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Double; begin a:=5.5; b:=5.25; WriteLn(Trunc(a+b)); end."#
        ),
        &["10"]
    );
}

#[test]
fn double_add_6() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Double; begin a:=6.5; b:=6.25; WriteLn(Trunc(a+b)); end."#
        ),
        &["12"]
    );
}

#[test]
fn double_add_7() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Double; begin a:=7.5; b:=7.25; WriteLn(Trunc(a+b)); end."#
        ),
        &["14"]
    );
}

#[test]
fn double_add_8() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Double; begin a:=8.5; b:=8.25; WriteLn(Trunc(a+b)); end."#
        ),
        &["16"]
    );
}

#[test]
fn double_add_9() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Double; begin a:=9.5; b:=9.25; WriteLn(Trunc(a+b)); end."#
        ),
        &["18"]
    );
}

#[test]
fn double_add_10() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Double; begin a:=10.5; b:=10.25; WriteLn(Trunc(a+b)); end."#
        ),
        &["20"]
    );
}

#[test]
fn double_add_11() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Double; begin a:=11.5; b:=11.25; WriteLn(Trunc(a+b)); end."#
        ),
        &["22"]
    );
}

#[test]
fn double_add_12() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Double; begin a:=12.5; b:=12.25; WriteLn(Trunc(a+b)); end."#
        ),
        &["24"]
    );
}

#[test]
fn double_add_13() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Double; begin a:=13.5; b:=13.25; WriteLn(Trunc(a+b)); end."#
        ),
        &["26"]
    );
}

#[test]
fn double_add_14() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Double; begin a:=14.5; b:=14.25; WriteLn(Trunc(a+b)); end."#
        ),
        &["28"]
    );
}

#[test]
fn extended_mul_1() {
    assert_eq!(
        run_pascal(r#"program T; var x:Extended; begin x:=1.0; WriteLn(Trunc(x*2)); end."#),
        &["2"]
    );
}

#[test]
fn extended_mul_2() {
    assert_eq!(
        run_pascal(r#"program T; var x:Extended; begin x:=2.0; WriteLn(Trunc(x*2)); end."#),
        &["4"]
    );
}

#[test]
fn extended_mul_3() {
    assert_eq!(
        run_pascal(r#"program T; var x:Extended; begin x:=3.0; WriteLn(Trunc(x*2)); end."#),
        &["6"]
    );
}

#[test]
fn extended_mul_4() {
    assert_eq!(
        run_pascal(r#"program T; var x:Extended; begin x:=4.0; WriteLn(Trunc(x*2)); end."#),
        &["8"]
    );
}

#[test]
fn extended_mul_5() {
    assert_eq!(
        run_pascal(r#"program T; var x:Extended; begin x:=5.0; WriteLn(Trunc(x*2)); end."#),
        &["10"]
    );
}

#[test]
fn extended_mul_6() {
    assert_eq!(
        run_pascal(r#"program T; var x:Extended; begin x:=6.0; WriteLn(Trunc(x*2)); end."#),
        &["12"]
    );
}

#[test]
fn extended_mul_7() {
    assert_eq!(
        run_pascal(r#"program T; var x:Extended; begin x:=7.0; WriteLn(Trunc(x*2)); end."#),
        &["14"]
    );
}

#[test]
fn extended_mul_8() {
    assert_eq!(
        run_pascal(r#"program T; var x:Extended; begin x:=8.0; WriteLn(Trunc(x*2)); end."#),
        &["16"]
    );
}

#[test]
fn extended_mul_9() {
    assert_eq!(
        run_pascal(r#"program T; var x:Extended; begin x:=9.0; WriteLn(Trunc(x*2)); end."#),
        &["18"]
    );
}

#[test]
fn extended_mul_10() {
    assert_eq!(
        run_pascal(r#"program T; var x:Extended; begin x:=10.0; WriteLn(Trunc(x*2)); end."#),
        &["20"]
    );
}

#[test]
fn extended_mul_11() {
    assert_eq!(
        run_pascal(r#"program T; var x:Extended; begin x:=11.0; WriteLn(Trunc(x*2)); end."#),
        &["22"]
    );
}

#[test]
fn extended_mul_12() {
    assert_eq!(
        run_pascal(r#"program T; var x:Extended; begin x:=12.0; WriteLn(Trunc(x*2)); end."#),
        &["24"]
    );
}

#[test]
fn extended_mul_13() {
    assert_eq!(
        run_pascal(r#"program T; var x:Extended; begin x:=13.0; WriteLn(Trunc(x*2)); end."#),
        &["26"]
    );
}

#[test]
fn extended_mul_14() {
    assert_eq!(
        run_pascal(r#"program T; var x:Extended; begin x:=14.0; WriteLn(Trunc(x*2)); end."#),
        &["28"]
    );
}

#[test]
fn currency_trunc_sum_1() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Currency; begin a:=1.10; b:=1.20; WriteLn(Trunc(a+b)); end."#
        ),
        &["2"]
    );
}

#[test]
fn currency_trunc_sum_2() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Currency; begin a:=2.10; b:=2.20; WriteLn(Trunc(a+b)); end."#
        ),
        &["4"]
    );
}

#[test]
fn currency_trunc_sum_3() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Currency; begin a:=3.10; b:=3.20; WriteLn(Trunc(a+b)); end."#
        ),
        &["6"]
    );
}

#[test]
fn currency_trunc_sum_4() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Currency; begin a:=4.10; b:=4.20; WriteLn(Trunc(a+b)); end."#
        ),
        &["8"]
    );
}

#[test]
fn currency_trunc_sum_5() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Currency; begin a:=5.10; b:=5.20; WriteLn(Trunc(a+b)); end."#
        ),
        &["10"]
    );
}

#[test]
fn currency_trunc_sum_6() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Currency; begin a:=6.10; b:=6.20; WriteLn(Trunc(a+b)); end."#
        ),
        &["12"]
    );
}

#[test]
fn currency_trunc_sum_7() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Currency; begin a:=7.10; b:=7.20; WriteLn(Trunc(a+b)); end."#
        ),
        &["14"]
    );
}

#[test]
fn currency_trunc_sum_8() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Currency; begin a:=8.10; b:=8.20; WriteLn(Trunc(a+b)); end."#
        ),
        &["16"]
    );
}

#[test]
fn currency_trunc_sum_9() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Currency; begin a:=9.10; b:=9.20; WriteLn(Trunc(a+b)); end."#
        ),
        &["18"]
    );
}

#[test]
fn currency_trunc_sum_10() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Currency; begin a:=10.10; b:=10.20; WriteLn(Trunc(a+b)); end."#
        ),
        &["20"]
    );
}

#[test]
fn currency_trunc_sum_11() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Currency; begin a:=11.10; b:=11.20; WriteLn(Trunc(a+b)); end."#
        ),
        &["22"]
    );
}

#[test]
fn currency_trunc_sum_12() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Currency; begin a:=12.10; b:=12.20; WriteLn(Trunc(a+b)); end."#
        ),
        &["24"]
    );
}

#[test]
fn currency_trunc_sum_13() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Currency; begin a:=13.10; b:=13.20; WriteLn(Trunc(a+b)); end."#
        ),
        &["26"]
    );
}

#[test]
fn currency_trunc_sum_14() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Currency; begin a:=14.10; b:=14.20; WriteLn(Trunc(a+b)); end."#
        ),
        &["28"]
    );
}
