/// Forward declarations and local nested procedures.
use super::helpers::run_pascal;

#[test]
fn forward_proc_1() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Alpha; forward; procedure Beta; begin WriteLn(1); end; procedure Alpha; begin Beta; end; begin Alpha; end."#
        ),
        &["1"]
    );
}

#[test]
fn forward_proc_2() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Alpha; forward; procedure Beta; begin WriteLn(2); end; procedure Alpha; begin Beta; end; begin Alpha; end."#
        ),
        &["2"]
    );
}

#[test]
fn forward_proc_3() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Alpha; forward; procedure Beta; begin WriteLn(3); end; procedure Alpha; begin Beta; end; begin Alpha; end."#
        ),
        &["3"]
    );
}

#[test]
fn forward_proc_4() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Alpha; forward; procedure Beta; begin WriteLn(4); end; procedure Alpha; begin Beta; end; begin Alpha; end."#
        ),
        &["4"]
    );
}

#[test]
fn forward_proc_5() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Alpha; forward; procedure Beta; begin WriteLn(5); end; procedure Alpha; begin Beta; end; begin Alpha; end."#
        ),
        &["5"]
    );
}

#[test]
fn forward_proc_6() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Alpha; forward; procedure Beta; begin WriteLn(6); end; procedure Alpha; begin Beta; end; begin Alpha; end."#
        ),
        &["6"]
    );
}

#[test]
fn forward_proc_7() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Alpha; forward; procedure Beta; begin WriteLn(7); end; procedure Alpha; begin Beta; end; begin Alpha; end."#
        ),
        &["7"]
    );
}

#[test]
fn forward_proc_8() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Alpha; forward; procedure Beta; begin WriteLn(8); end; procedure Alpha; begin Beta; end; begin Alpha; end."#
        ),
        &["8"]
    );
}

#[test]
fn forward_proc_9() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Alpha; forward; procedure Beta; begin WriteLn(9); end; procedure Alpha; begin Beta; end; begin Alpha; end."#
        ),
        &["9"]
    );
}

#[test]
fn forward_proc_10() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Alpha; forward; procedure Beta; begin WriteLn(10); end; procedure Alpha; begin Beta; end; begin Alpha; end."#
        ),
        &["10"]
    );
}

#[test]
fn forward_proc_11() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Alpha; forward; procedure Beta; begin WriteLn(11); end; procedure Alpha; begin Beta; end; begin Alpha; end."#
        ),
        &["11"]
    );
}

#[test]
fn forward_proc_12() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Alpha; forward; procedure Beta; begin WriteLn(12); end; procedure Alpha; begin Beta; end; begin Alpha; end."#
        ),
        &["12"]
    );
}

#[test]
fn forward_proc_13() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Alpha; forward; procedure Beta; begin WriteLn(13); end; procedure Alpha; begin Beta; end; begin Alpha; end."#
        ),
        &["13"]
    );
}

#[test]
fn forward_proc_14() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Alpha; forward; procedure Beta; begin WriteLn(14); end; procedure Alpha; begin Beta; end; begin Alpha; end."#
        ),
        &["14"]
    );
}

#[test]
fn forward_proc_15() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Alpha; forward; procedure Beta; begin WriteLn(15); end; procedure Alpha; begin Beta; end; begin Alpha; end."#
        ),
        &["15"]
    );
}

#[test]
fn forward_proc_16() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Alpha; forward; procedure Beta; begin WriteLn(16); end; procedure Alpha; begin Beta; end; begin Alpha; end."#
        ),
        &["16"]
    );
}

#[test]
fn forward_proc_17() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Alpha; forward; procedure Beta; begin WriteLn(17); end; procedure Alpha; begin Beta; end; begin Alpha; end."#
        ),
        &["17"]
    );
}

#[test]
fn forward_proc_18() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Alpha; forward; procedure Beta; begin WriteLn(18); end; procedure Alpha; begin Beta; end; begin Alpha; end."#
        ),
        &["18"]
    );
}

#[test]
fn forward_proc_19() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Alpha; forward; procedure Beta; begin WriteLn(19); end; procedure Alpha; begin Beta; end; begin Alpha; end."#
        ),
        &["19"]
    );
}

#[test]
fn forward_proc_20() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Alpha; forward; procedure Beta; begin WriteLn(20); end; procedure Alpha; begin Beta; end; begin Alpha; end."#
        ),
        &["20"]
    );
}

#[test]
fn forward_proc_21() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Alpha; forward; procedure Beta; begin WriteLn(21); end; procedure Alpha; begin Beta; end; begin Alpha; end."#
        ),
        &["21"]
    );
}

#[test]
fn forward_func_1() {
    assert_eq!(
        run_pascal(
            r#"program T; function DoubleIt(n:Integer):Integer; forward; function TripleIt(n:Integer):Integer; begin Result:=DoubleIt(n)+n; end; function DoubleIt(n:Integer):Integer; begin Result:=n*2; end; begin WriteLn(TripleIt(1)); end."#
        ),
        &["3"]
    );
}

#[test]
fn forward_func_2() {
    assert_eq!(
        run_pascal(
            r#"program T; function DoubleIt(n:Integer):Integer; forward; function TripleIt(n:Integer):Integer; begin Result:=DoubleIt(n)+n; end; function DoubleIt(n:Integer):Integer; begin Result:=n*2; end; begin WriteLn(TripleIt(2)); end."#
        ),
        &["6"]
    );
}

#[test]
fn forward_func_3() {
    assert_eq!(
        run_pascal(
            r#"program T; function DoubleIt(n:Integer):Integer; forward; function TripleIt(n:Integer):Integer; begin Result:=DoubleIt(n)+n; end; function DoubleIt(n:Integer):Integer; begin Result:=n*2; end; begin WriteLn(TripleIt(3)); end."#
        ),
        &["9"]
    );
}

#[test]
fn forward_func_4() {
    assert_eq!(
        run_pascal(
            r#"program T; function DoubleIt(n:Integer):Integer; forward; function TripleIt(n:Integer):Integer; begin Result:=DoubleIt(n)+n; end; function DoubleIt(n:Integer):Integer; begin Result:=n*2; end; begin WriteLn(TripleIt(4)); end."#
        ),
        &["12"]
    );
}

#[test]
fn forward_func_5() {
    assert_eq!(
        run_pascal(
            r#"program T; function DoubleIt(n:Integer):Integer; forward; function TripleIt(n:Integer):Integer; begin Result:=DoubleIt(n)+n; end; function DoubleIt(n:Integer):Integer; begin Result:=n*2; end; begin WriteLn(TripleIt(5)); end."#
        ),
        &["15"]
    );
}

#[test]
fn forward_func_6() {
    assert_eq!(
        run_pascal(
            r#"program T; function DoubleIt(n:Integer):Integer; forward; function TripleIt(n:Integer):Integer; begin Result:=DoubleIt(n)+n; end; function DoubleIt(n:Integer):Integer; begin Result:=n*2; end; begin WriteLn(TripleIt(6)); end."#
        ),
        &["18"]
    );
}

#[test]
fn forward_func_7() {
    assert_eq!(
        run_pascal(
            r#"program T; function DoubleIt(n:Integer):Integer; forward; function TripleIt(n:Integer):Integer; begin Result:=DoubleIt(n)+n; end; function DoubleIt(n:Integer):Integer; begin Result:=n*2; end; begin WriteLn(TripleIt(7)); end."#
        ),
        &["21"]
    );
}

#[test]
fn forward_func_8() {
    assert_eq!(
        run_pascal(
            r#"program T; function DoubleIt(n:Integer):Integer; forward; function TripleIt(n:Integer):Integer; begin Result:=DoubleIt(n)+n; end; function DoubleIt(n:Integer):Integer; begin Result:=n*2; end; begin WriteLn(TripleIt(8)); end."#
        ),
        &["24"]
    );
}

#[test]
fn forward_func_9() {
    assert_eq!(
        run_pascal(
            r#"program T; function DoubleIt(n:Integer):Integer; forward; function TripleIt(n:Integer):Integer; begin Result:=DoubleIt(n)+n; end; function DoubleIt(n:Integer):Integer; begin Result:=n*2; end; begin WriteLn(TripleIt(9)); end."#
        ),
        &["27"]
    );
}

#[test]
fn forward_func_10() {
    assert_eq!(
        run_pascal(
            r#"program T; function DoubleIt(n:Integer):Integer; forward; function TripleIt(n:Integer):Integer; begin Result:=DoubleIt(n)+n; end; function DoubleIt(n:Integer):Integer; begin Result:=n*2; end; begin WriteLn(TripleIt(10)); end."#
        ),
        &["30"]
    );
}

#[test]
fn forward_func_11() {
    assert_eq!(
        run_pascal(
            r#"program T; function DoubleIt(n:Integer):Integer; forward; function TripleIt(n:Integer):Integer; begin Result:=DoubleIt(n)+n; end; function DoubleIt(n:Integer):Integer; begin Result:=n*2; end; begin WriteLn(TripleIt(11)); end."#
        ),
        &["33"]
    );
}

#[test]
fn forward_func_12() {
    assert_eq!(
        run_pascal(
            r#"program T; function DoubleIt(n:Integer):Integer; forward; function TripleIt(n:Integer):Integer; begin Result:=DoubleIt(n)+n; end; function DoubleIt(n:Integer):Integer; begin Result:=n*2; end; begin WriteLn(TripleIt(12)); end."#
        ),
        &["36"]
    );
}

#[test]
fn forward_func_13() {
    assert_eq!(
        run_pascal(
            r#"program T; function DoubleIt(n:Integer):Integer; forward; function TripleIt(n:Integer):Integer; begin Result:=DoubleIt(n)+n; end; function DoubleIt(n:Integer):Integer; begin Result:=n*2; end; begin WriteLn(TripleIt(13)); end."#
        ),
        &["39"]
    );
}

#[test]
fn forward_func_14() {
    assert_eq!(
        run_pascal(
            r#"program T; function DoubleIt(n:Integer):Integer; forward; function TripleIt(n:Integer):Integer; begin Result:=DoubleIt(n)+n; end; function DoubleIt(n:Integer):Integer; begin Result:=n*2; end; begin WriteLn(TripleIt(14)); end."#
        ),
        &["42"]
    );
}

#[test]
fn forward_func_15() {
    assert_eq!(
        run_pascal(
            r#"program T; function DoubleIt(n:Integer):Integer; forward; function TripleIt(n:Integer):Integer; begin Result:=DoubleIt(n)+n; end; function DoubleIt(n:Integer):Integer; begin Result:=n*2; end; begin WriteLn(TripleIt(15)); end."#
        ),
        &["45"]
    );
}

#[test]
fn forward_func_16() {
    assert_eq!(
        run_pascal(
            r#"program T; function DoubleIt(n:Integer):Integer; forward; function TripleIt(n:Integer):Integer; begin Result:=DoubleIt(n)+n; end; function DoubleIt(n:Integer):Integer; begin Result:=n*2; end; begin WriteLn(TripleIt(16)); end."#
        ),
        &["48"]
    );
}

#[test]
fn forward_func_17() {
    assert_eq!(
        run_pascal(
            r#"program T; function DoubleIt(n:Integer):Integer; forward; function TripleIt(n:Integer):Integer; begin Result:=DoubleIt(n)+n; end; function DoubleIt(n:Integer):Integer; begin Result:=n*2; end; begin WriteLn(TripleIt(17)); end."#
        ),
        &["51"]
    );
}

#[test]
fn forward_func_18() {
    assert_eq!(
        run_pascal(
            r#"program T; function DoubleIt(n:Integer):Integer; forward; function TripleIt(n:Integer):Integer; begin Result:=DoubleIt(n)+n; end; function DoubleIt(n:Integer):Integer; begin Result:=n*2; end; begin WriteLn(TripleIt(18)); end."#
        ),
        &["54"]
    );
}

#[test]
fn forward_func_19() {
    assert_eq!(
        run_pascal(
            r#"program T; function DoubleIt(n:Integer):Integer; forward; function TripleIt(n:Integer):Integer; begin Result:=DoubleIt(n)+n; end; function DoubleIt(n:Integer):Integer; begin Result:=n*2; end; begin WriteLn(TripleIt(19)); end."#
        ),
        &["57"]
    );
}

#[test]
fn forward_func_20() {
    assert_eq!(
        run_pascal(
            r#"program T; function DoubleIt(n:Integer):Integer; forward; function TripleIt(n:Integer):Integer; begin Result:=DoubleIt(n)+n; end; function DoubleIt(n:Integer):Integer; begin Result:=n*2; end; begin WriteLn(TripleIt(20)); end."#
        ),
        &["60"]
    );
}

#[test]
fn forward_func_21() {
    assert_eq!(
        run_pascal(
            r#"program T; function DoubleIt(n:Integer):Integer; forward; function TripleIt(n:Integer):Integer; begin Result:=DoubleIt(n)+n; end; function DoubleIt(n:Integer):Integer; begin Result:=n*2; end; begin WriteLn(TripleIt(21)); end."#
        ),
        &["63"]
    );
}

