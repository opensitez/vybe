/// Additional anonymous method and closure patterns.
use super::helpers::run_pascal;

#[test]
fn anonx_proc_1() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:procedure; begin p:=procedure begin WriteLn('anon1'); end; p; end."#
        ),
        &["anon1"]
    );
}

#[test]
fn anonx_func_2() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:function:Integer; begin f:=function:Integer begin Result:=2; end; WriteLn(f()); end."#
        ),
        &["2"]
    );
}

#[test]
fn anonx_param_3() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:function(x:Integer):Integer; begin f:=function(x:Integer):Integer begin Result:=x+3; end; WriteLn(f(3)); end."#
        ),
        &["6"]
    );
}

#[test]
fn anonx_callback_4() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Run(p:procedure); begin p; end; begin Run(procedure begin WriteLn('cb4'); end); end."#
        ),
        &["cb4"]
    );
}

#[test]
fn anonx_apply_5() {
    assert_eq!(
        run_pascal(
            r#"program T; function Apply(fn:function(x:Integer):Integer; v:Integer):Integer; begin Result:=fn(v); end; begin WriteLn(Apply(function(x:Integer):Integer begin Result:=x*5; end,5)); end."#
        ),
        &["25"]
    );
}

#[test]
fn anonx_assign_6() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:procedure; begin a:=procedure begin WriteLn('copy6'); end; b:=a; b; end."#
        ),
        &["copy6"]
    );
}

#[test]
fn anonx_proc_7() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:procedure; begin p:=procedure begin WriteLn('anon7'); end; p; end."#
        ),
        &["anon7"]
    );
}

#[test]
fn anonx_func_8() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:function:Integer; begin f:=function:Integer begin Result:=8; end; WriteLn(f()); end."#
        ),
        &["8"]
    );
}

#[test]
fn anonx_param_9() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:function(x:Integer):Integer; begin f:=function(x:Integer):Integer begin Result:=x+9; end; WriteLn(f(9)); end."#
        ),
        &["18"]
    );
}

#[test]
fn anonx_callback_10() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Run(p:procedure); begin p; end; begin Run(procedure begin WriteLn('cb10'); end); end."#
        ),
        &["cb10"]
    );
}

#[test]
fn anonx_apply_11() {
    assert_eq!(
        run_pascal(
            r#"program T; function Apply(fn:function(x:Integer):Integer; v:Integer):Integer; begin Result:=fn(v); end; begin WriteLn(Apply(function(x:Integer):Integer begin Result:=x*11; end,11)); end."#
        ),
        &["121"]
    );
}

#[test]
fn anonx_assign_12() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:procedure; begin a:=procedure begin WriteLn('copy12'); end; b:=a; b; end."#
        ),
        &["copy12"]
    );
}

#[test]
fn anonx_proc_13() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:procedure; begin p:=procedure begin WriteLn('anon13'); end; p; end."#
        ),
        &["anon13"]
    );
}

#[test]
fn anonx_func_14() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:function:Integer; begin f:=function:Integer begin Result:=14; end; WriteLn(f()); end."#
        ),
        &["14"]
    );
}

#[test]
fn anonx_param_15() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:function(x:Integer):Integer; begin f:=function(x:Integer):Integer begin Result:=x+15; end; WriteLn(f(15)); end."#
        ),
        &["30"]
    );
}

#[test]
fn anonx_callback_16() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Run(p:procedure); begin p; end; begin Run(procedure begin WriteLn('cb16'); end); end."#
        ),
        &["cb16"]
    );
}

#[test]
fn anonx_apply_17() {
    assert_eq!(
        run_pascal(
            r#"program T; function Apply(fn:function(x:Integer):Integer; v:Integer):Integer; begin Result:=fn(v); end; begin WriteLn(Apply(function(x:Integer):Integer begin Result:=x*17; end,17)); end."#
        ),
        &["289"]
    );
}

#[test]
fn anonx_assign_18() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:procedure; begin a:=procedure begin WriteLn('copy18'); end; b:=a; b; end."#
        ),
        &["copy18"]
    );
}

#[test]
fn anonx_proc_19() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:procedure; begin p:=procedure begin WriteLn('anon19'); end; p; end."#
        ),
        &["anon19"]
    );
}

#[test]
fn anonx_func_20() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:function:Integer; begin f:=function:Integer begin Result:=20; end; WriteLn(f()); end."#
        ),
        &["20"]
    );
}

#[test]
fn anonx_param_21() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:function(x:Integer):Integer; begin f:=function(x:Integer):Integer begin Result:=x+21; end; WriteLn(f(21)); end."#
        ),
        &["42"]
    );
}

#[test]
fn anonx_callback_22() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Run(p:procedure); begin p; end; begin Run(procedure begin WriteLn('cb22'); end); end."#
        ),
        &["cb22"]
    );
}

#[test]
fn anonx_apply_23() {
    assert_eq!(
        run_pascal(
            r#"program T; function Apply(fn:function(x:Integer):Integer; v:Integer):Integer; begin Result:=fn(v); end; begin WriteLn(Apply(function(x:Integer):Integer begin Result:=x*23; end,23)); end."#
        ),
        &["529"]
    );
}

#[test]
fn anonx_assign_24() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:procedure; begin a:=procedure begin WriteLn('copy24'); end; b:=a; b; end."#
        ),
        &["copy24"]
    );
}

#[test]
fn anonx_proc_25() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:procedure; begin p:=procedure begin WriteLn('anon25'); end; p; end."#
        ),
        &["anon25"]
    );
}

#[test]
fn anonx_func_26() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:function:Integer; begin f:=function:Integer begin Result:=26; end; WriteLn(f()); end."#
        ),
        &["26"]
    );
}

#[test]
fn anonx_param_27() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:function(x:Integer):Integer; begin f:=function(x:Integer):Integer begin Result:=x+27; end; WriteLn(f(27)); end."#
        ),
        &["54"]
    );
}

#[test]
fn anonx_callback_28() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Run(p:procedure); begin p; end; begin Run(procedure begin WriteLn('cb28'); end); end."#
        ),
        &["cb28"]
    );
}

#[test]
fn anonx_apply_29() {
    assert_eq!(
        run_pascal(
            r#"program T; function Apply(fn:function(x:Integer):Integer; v:Integer):Integer; begin Result:=fn(v); end; begin WriteLn(Apply(function(x:Integer):Integer begin Result:=x*29; end,29)); end."#
        ),
        &["841"]
    );
}

#[test]
fn anonx_assign_30() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:procedure; begin a:=procedure begin WriteLn('copy30'); end; b:=a; b; end."#
        ),
        &["copy30"]
    );
}

#[test]
fn anonx_proc_31() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:procedure; begin p:=procedure begin WriteLn('anon31'); end; p; end."#
        ),
        &["anon31"]
    );
}

#[test]
fn anonx_func_32() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:function:Integer; begin f:=function:Integer begin Result:=32; end; WriteLn(f()); end."#
        ),
        &["32"]
    );
}

#[test]
fn anonx_param_33() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:function(x:Integer):Integer; begin f:=function(x:Integer):Integer begin Result:=x+33; end; WriteLn(f(33)); end."#
        ),
        &["66"]
    );
}

#[test]
fn anonx_callback_34() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Run(p:procedure); begin p; end; begin Run(procedure begin WriteLn('cb34'); end); end."#
        ),
        &["cb34"]
    );
}

#[test]
fn anonx_apply_35() {
    assert_eq!(
        run_pascal(
            r#"program T; function Apply(fn:function(x:Integer):Integer; v:Integer):Integer; begin Result:=fn(v); end; begin WriteLn(Apply(function(x:Integer):Integer begin Result:=x*35; end,35)); end."#
        ),
        &["1225"]
    );
}

#[test]
fn anonx_assign_36() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:procedure; begin a:=procedure begin WriteLn('copy36'); end; b:=a; b; end."#
        ),
        &["copy36"]
    );
}

#[test]
fn anonx_proc_37() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:procedure; begin p:=procedure begin WriteLn('anon37'); end; p; end."#
        ),
        &["anon37"]
    );
}

#[test]
fn anonx_func_38() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:function:Integer; begin f:=function:Integer begin Result:=38; end; WriteLn(f()); end."#
        ),
        &["38"]
    );
}

#[test]
fn anonx_param_39() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:function(x:Integer):Integer; begin f:=function(x:Integer):Integer begin Result:=x+39; end; WriteLn(f(39)); end."#
        ),
        &["78"]
    );
}

#[test]
fn anonx_callback_40() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Run(p:procedure); begin p; end; begin Run(procedure begin WriteLn('cb40'); end); end."#
        ),
        &["cb40"]
    );
}

#[test]
fn anonx_apply_41() {
    assert_eq!(
        run_pascal(
            r#"program T; function Apply(fn:function(x:Integer):Integer; v:Integer):Integer; begin Result:=fn(v); end; begin WriteLn(Apply(function(x:Integer):Integer begin Result:=x*41; end,41)); end."#
        ),
        &["1681"]
    );
}

#[test]
fn anonx_assign_42() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:procedure; begin a:=procedure begin WriteLn('copy42'); end; b:=a; b; end."#
        ),
        &["copy42"]
    );
}
