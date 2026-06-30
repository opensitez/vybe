/// Procedural type assignments and dispatch.
use super::helpers::run_pascal;

#[test]
fn proctypex_fn_var_1() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:function(x:Integer):Integer; begin f:=function(x:Integer):Integer begin Result:=x+1; end; WriteLn(f(1)); end."#
        ),
        &["2"]
    );
}

#[test]
fn proctypex_proc_var_2() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:procedure; begin p:=procedure begin WriteLn('p2'); end; p; end."#
        ),
        &["p2"]
    );
}

#[test]
fn proctypex_twice_3() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Twice(f:function(x:Integer):Integer); begin WriteLn(f(3)); WriteLn(f(3)); end; begin Twice(function(x:Integer):Integer begin Result:=x*2; end); end."#
        ),
        &["6", "6"]
    );
}

#[test]
fn proctypex_factory_4() {
    assert_eq!(
        run_pascal(
            r#"program T; function Make(k:Integer):function(x:Integer):Integer; begin Result:=function(x:Integer):Integer begin Result:=x+k; end; end; var f:function(x:Integer):Integer; begin f:=Make(4); WriteLn(f(0)); end."#
        ),
        &["4"]
    );
}

#[test]
fn proctypex_capture_5() {
    assert_eq!(
        run_pascal(
            r#"program T; var base:Integer; f:function(x:Integer):Integer; begin base:=5; f:=function(x:Integer):Integer begin Result:=base+x; end; WriteLn(f(1)); end."#
        ),
        &["6"]
    );
}

#[test]
fn proctypex_reassign_6() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:function(x:Integer):Integer; begin f:=function(x:Integer):Integer begin Result:=x; end; WriteLn(f(6)); f:=function(x:Integer):Integer begin Result:=x+10; end; WriteLn(f(6)); end."#
        ),
        &["6", "16"]
    );
}

#[test]
fn proctypex_fn_var_7() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:function(x:Integer):Integer; begin f:=function(x:Integer):Integer begin Result:=x+7; end; WriteLn(f(1)); end."#
        ),
        &["8"]
    );
}

#[test]
fn proctypex_proc_var_8() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:procedure; begin p:=procedure begin WriteLn('p8'); end; p; end."#
        ),
        &["p8"]
    );
}

#[test]
fn proctypex_twice_9() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Twice(f:function(x:Integer):Integer); begin WriteLn(f(9)); WriteLn(f(9)); end; begin Twice(function(x:Integer):Integer begin Result:=x*2; end); end."#
        ),
        &["18", "18"]
    );
}

#[test]
fn proctypex_factory_10() {
    assert_eq!(
        run_pascal(
            r#"program T; function Make(k:Integer):function(x:Integer):Integer; begin Result:=function(x:Integer):Integer begin Result:=x+k; end; end; var f:function(x:Integer):Integer; begin f:=Make(10); WriteLn(f(0)); end."#
        ),
        &["10"]
    );
}

#[test]
fn proctypex_capture_11() {
    assert_eq!(
        run_pascal(
            r#"program T; var base:Integer; f:function(x:Integer):Integer; begin base:=11; f:=function(x:Integer):Integer begin Result:=base+x; end; WriteLn(f(1)); end."#
        ),
        &["12"]
    );
}

#[test]
fn proctypex_reassign_12() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:function(x:Integer):Integer; begin f:=function(x:Integer):Integer begin Result:=x; end; WriteLn(f(12)); f:=function(x:Integer):Integer begin Result:=x+10; end; WriteLn(f(12)); end."#
        ),
        &["12", "22"]
    );
}

#[test]
fn proctypex_fn_var_13() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:function(x:Integer):Integer; begin f:=function(x:Integer):Integer begin Result:=x+13; end; WriteLn(f(1)); end."#
        ),
        &["14"]
    );
}

#[test]
fn proctypex_proc_var_14() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:procedure; begin p:=procedure begin WriteLn('p14'); end; p; end."#
        ),
        &["p14"]
    );
}

#[test]
fn proctypex_twice_15() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Twice(f:function(x:Integer):Integer); begin WriteLn(f(15)); WriteLn(f(15)); end; begin Twice(function(x:Integer):Integer begin Result:=x*2; end); end."#
        ),
        &["30", "30"]
    );
}

#[test]
fn proctypex_factory_16() {
    assert_eq!(
        run_pascal(
            r#"program T; function Make(k:Integer):function(x:Integer):Integer; begin Result:=function(x:Integer):Integer begin Result:=x+k; end; end; var f:function(x:Integer):Integer; begin f:=Make(16); WriteLn(f(0)); end."#
        ),
        &["16"]
    );
}

#[test]
fn proctypex_capture_17() {
    assert_eq!(
        run_pascal(
            r#"program T; var base:Integer; f:function(x:Integer):Integer; begin base:=17; f:=function(x:Integer):Integer begin Result:=base+x; end; WriteLn(f(1)); end."#
        ),
        &["18"]
    );
}

#[test]
fn proctypex_reassign_18() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:function(x:Integer):Integer; begin f:=function(x:Integer):Integer begin Result:=x; end; WriteLn(f(18)); f:=function(x:Integer):Integer begin Result:=x+10; end; WriteLn(f(18)); end."#
        ),
        &["18", "28"]
    );
}

#[test]
fn proctypex_fn_var_19() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:function(x:Integer):Integer; begin f:=function(x:Integer):Integer begin Result:=x+19; end; WriteLn(f(1)); end."#
        ),
        &["20"]
    );
}

#[test]
fn proctypex_proc_var_20() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:procedure; begin p:=procedure begin WriteLn('p20'); end; p; end."#
        ),
        &["p20"]
    );
}

#[test]
fn proctypex_twice_21() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Twice(f:function(x:Integer):Integer); begin WriteLn(f(21)); WriteLn(f(21)); end; begin Twice(function(x:Integer):Integer begin Result:=x*2; end); end."#
        ),
        &["42", "42"]
    );
}

#[test]
fn proctypex_factory_22() {
    assert_eq!(
        run_pascal(
            r#"program T; function Make(k:Integer):function(x:Integer):Integer; begin Result:=function(x:Integer):Integer begin Result:=x+k; end; end; var f:function(x:Integer):Integer; begin f:=Make(22); WriteLn(f(0)); end."#
        ),
        &["22"]
    );
}

#[test]
fn proctypex_capture_23() {
    assert_eq!(
        run_pascal(
            r#"program T; var base:Integer; f:function(x:Integer):Integer; begin base:=23; f:=function(x:Integer):Integer begin Result:=base+x; end; WriteLn(f(1)); end."#
        ),
        &["24"]
    );
}

#[test]
fn proctypex_reassign_24() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:function(x:Integer):Integer; begin f:=function(x:Integer):Integer begin Result:=x; end; WriteLn(f(24)); f:=function(x:Integer):Integer begin Result:=x+10; end; WriteLn(f(24)); end."#
        ),
        &["24", "34"]
    );
}

#[test]
fn proctypex_fn_var_25() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:function(x:Integer):Integer; begin f:=function(x:Integer):Integer begin Result:=x+25; end; WriteLn(f(1)); end."#
        ),
        &["26"]
    );
}

#[test]
fn proctypex_proc_var_26() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:procedure; begin p:=procedure begin WriteLn('p26'); end; p; end."#
        ),
        &["p26"]
    );
}

#[test]
fn proctypex_twice_27() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Twice(f:function(x:Integer):Integer); begin WriteLn(f(27)); WriteLn(f(27)); end; begin Twice(function(x:Integer):Integer begin Result:=x*2; end); end."#
        ),
        &["54", "54"]
    );
}

#[test]
fn proctypex_factory_28() {
    assert_eq!(
        run_pascal(
            r#"program T; function Make(k:Integer):function(x:Integer):Integer; begin Result:=function(x:Integer):Integer begin Result:=x+k; end; end; var f:function(x:Integer):Integer; begin f:=Make(28); WriteLn(f(0)); end."#
        ),
        &["28"]
    );
}

#[test]
fn proctypex_capture_29() {
    assert_eq!(
        run_pascal(
            r#"program T; var base:Integer; f:function(x:Integer):Integer; begin base:=29; f:=function(x:Integer):Integer begin Result:=base+x; end; WriteLn(f(1)); end."#
        ),
        &["30"]
    );
}

#[test]
fn proctypex_reassign_30() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:function(x:Integer):Integer; begin f:=function(x:Integer):Integer begin Result:=x; end; WriteLn(f(30)); f:=function(x:Integer):Integer begin Result:=x+10; end; WriteLn(f(30)); end."#
        ),
        &["30", "40"]
    );
}

#[test]
fn proctypex_fn_var_31() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:function(x:Integer):Integer; begin f:=function(x:Integer):Integer begin Result:=x+31; end; WriteLn(f(1)); end."#
        ),
        &["32"]
    );
}

#[test]
fn proctypex_proc_var_32() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:procedure; begin p:=procedure begin WriteLn('p32'); end; p; end."#
        ),
        &["p32"]
    );
}

#[test]
fn proctypex_twice_33() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Twice(f:function(x:Integer):Integer); begin WriteLn(f(33)); WriteLn(f(33)); end; begin Twice(function(x:Integer):Integer begin Result:=x*2; end); end."#
        ),
        &["66", "66"]
    );
}

#[test]
fn proctypex_factory_34() {
    assert_eq!(
        run_pascal(
            r#"program T; function Make(k:Integer):function(x:Integer):Integer; begin Result:=function(x:Integer):Integer begin Result:=x+k; end; end; var f:function(x:Integer):Integer; begin f:=Make(34); WriteLn(f(0)); end."#
        ),
        &["34"]
    );
}

#[test]
fn proctypex_capture_35() {
    assert_eq!(
        run_pascal(
            r#"program T; var base:Integer; f:function(x:Integer):Integer; begin base:=35; f:=function(x:Integer):Integer begin Result:=base+x; end; WriteLn(f(1)); end."#
        ),
        &["36"]
    );
}

#[test]
fn proctypex_reassign_36() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:function(x:Integer):Integer; begin f:=function(x:Integer):Integer begin Result:=x; end; WriteLn(f(36)); f:=function(x:Integer):Integer begin Result:=x+10; end; WriteLn(f(36)); end."#
        ),
        &["36", "46"]
    );
}

#[test]
fn proctypex_fn_var_37() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:function(x:Integer):Integer; begin f:=function(x:Integer):Integer begin Result:=x+37; end; WriteLn(f(1)); end."#
        ),
        &["38"]
    );
}

#[test]
fn proctypex_proc_var_38() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:procedure; begin p:=procedure begin WriteLn('p38'); end; p; end."#
        ),
        &["p38"]
    );
}

#[test]
fn proctypex_twice_39() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Twice(f:function(x:Integer):Integer); begin WriteLn(f(39)); WriteLn(f(39)); end; begin Twice(function(x:Integer):Integer begin Result:=x*2; end); end."#
        ),
        &["78", "78"]
    );
}

#[test]
fn proctypex_factory_40() {
    assert_eq!(
        run_pascal(
            r#"program T; function Make(k:Integer):function(x:Integer):Integer; begin Result:=function(x:Integer):Integer begin Result:=x+k; end; end; var f:function(x:Integer):Integer; begin f:=Make(40); WriteLn(f(0)); end."#
        ),
        &["40"]
    );
}

#[test]
fn proctypex_capture_41() {
    assert_eq!(
        run_pascal(
            r#"program T; var base:Integer; f:function(x:Integer):Integer; begin base:=41; f:=function(x:Integer):Integer begin Result:=base+x; end; WriteLn(f(1)); end."#
        ),
        &["42"]
    );
}

#[test]
fn proctypex_reassign_42() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:function(x:Integer):Integer; begin f:=function(x:Integer):Integer begin Result:=x; end; WriteLn(f(42)); f:=function(x:Integer):Integer begin Result:=x+10; end; WriteLn(f(42)); end."#
        ),
        &["42", "52"]
    );
}
