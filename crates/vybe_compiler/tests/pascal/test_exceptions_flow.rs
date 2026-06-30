/// Try/except/finally flow combinations.
use super::helpers::run_pascal;

#[test]
fn excflow_try_ok_1() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try WriteLn('try1'); except WriteLn('ex'); end; end."#
        ),
        &["try1"]
    );
}

#[test]
fn excflow_catch_raise_2() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try raise Exception.Create('boom2'); except WriteLn('got2'); end; end."#
        ),
        &["got2"]
    );
}

#[test]
fn excflow_finally_3() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try WriteLn('x3'); finally WriteLn('fin3'); end; end."#
        ),
        &["x3", "fin3"]
    );
}

#[test]
fn excflow_except_finally_4() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try try raise Exception.Create('e4'); except WriteLn('c4'); end; finally WriteLn('f4'); end; end."#
        ),
        &["c4", "f4"]
    );
}

#[test]
fn excflow_typed_5() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try raise ERangeError.Create('rng5'); except on E:ERangeError do WriteLn('range5'); on E:Exception do WriteLn('other'); end; end."#
        ),
        &["range5"]
    );
}

#[test]
fn excflow_div_zero_6() {
    assert_eq!(
        run_pascal(
            r#"program T; var v:Integer; begin v:=6; try v:=v div 0; except v:=-1; end; WriteLn(v); end."#
        ),
        &["-1"]
    );
}

#[test]
fn excflow_nested_7() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try try try raise Exception.Create('deep7'); except WriteLn('L37'); end; except WriteLn('L2'); end; except WriteLn('L1'); end; end."#
        ),
        &["L37"]
    );
}

#[test]
fn excflow_try_ok_8() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try WriteLn('try8'); except WriteLn('ex'); end; end."#
        ),
        &["try8"]
    );
}

#[test]
fn excflow_catch_raise_9() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try raise Exception.Create('boom9'); except WriteLn('got9'); end; end."#
        ),
        &["got9"]
    );
}

#[test]
fn excflow_finally_10() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try WriteLn('x10'); finally WriteLn('fin10'); end; end."#
        ),
        &["x10", "fin10"]
    );
}

#[test]
fn excflow_except_finally_11() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try try raise Exception.Create('e11'); except WriteLn('c11'); end; finally WriteLn('f11'); end; end."#
        ),
        &["c11", "f11"]
    );
}

#[test]
fn excflow_typed_12() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try raise ERangeError.Create('rng12'); except on E:ERangeError do WriteLn('range12'); on E:Exception do WriteLn('other'); end; end."#
        ),
        &["range12"]
    );
}

#[test]
fn excflow_div_zero_13() {
    assert_eq!(
        run_pascal(
            r#"program T; var v:Integer; begin v:=13; try v:=v div 0; except v:=-1; end; WriteLn(v); end."#
        ),
        &["-1"]
    );
}

#[test]
fn excflow_nested_14() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try try try raise Exception.Create('deep14'); except WriteLn('L314'); end; except WriteLn('L2'); end; except WriteLn('L1'); end; end."#
        ),
        &["L314"]
    );
}

#[test]
fn excflow_try_ok_15() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try WriteLn('try15'); except WriteLn('ex'); end; end."#
        ),
        &["try15"]
    );
}

#[test]
fn excflow_catch_raise_16() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try raise Exception.Create('boom16'); except WriteLn('got16'); end; end."#
        ),
        &["got16"]
    );
}

#[test]
fn excflow_finally_17() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try WriteLn('x17'); finally WriteLn('fin17'); end; end."#
        ),
        &["x17", "fin17"]
    );
}

#[test]
fn excflow_except_finally_18() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try try raise Exception.Create('e18'); except WriteLn('c18'); end; finally WriteLn('f18'); end; end."#
        ),
        &["c18", "f18"]
    );
}

#[test]
fn excflow_typed_19() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try raise ERangeError.Create('rng19'); except on E:ERangeError do WriteLn('range19'); on E:Exception do WriteLn('other'); end; end."#
        ),
        &["range19"]
    );
}

#[test]
fn excflow_div_zero_20() {
    assert_eq!(
        run_pascal(
            r#"program T; var v:Integer; begin v:=20; try v:=v div 0; except v:=-1; end; WriteLn(v); end."#
        ),
        &["-1"]
    );
}

#[test]
fn excflow_nested_21() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try try try raise Exception.Create('deep21'); except WriteLn('L321'); end; except WriteLn('L2'); end; except WriteLn('L1'); end; end."#
        ),
        &["L321"]
    );
}

#[test]
fn excflow_try_ok_22() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try WriteLn('try22'); except WriteLn('ex'); end; end."#
        ),
        &["try22"]
    );
}

#[test]
fn excflow_catch_raise_23() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try raise Exception.Create('boom23'); except WriteLn('got23'); end; end."#
        ),
        &["got23"]
    );
}

#[test]
fn excflow_finally_24() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try WriteLn('x24'); finally WriteLn('fin24'); end; end."#
        ),
        &["x24", "fin24"]
    );
}

#[test]
fn excflow_except_finally_25() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try try raise Exception.Create('e25'); except WriteLn('c25'); end; finally WriteLn('f25'); end; end."#
        ),
        &["c25", "f25"]
    );
}

#[test]
fn excflow_typed_26() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try raise ERangeError.Create('rng26'); except on E:ERangeError do WriteLn('range26'); on E:Exception do WriteLn('other'); end; end."#
        ),
        &["range26"]
    );
}

#[test]
fn excflow_div_zero_27() {
    assert_eq!(
        run_pascal(
            r#"program T; var v:Integer; begin v:=27; try v:=v div 0; except v:=-1; end; WriteLn(v); end."#
        ),
        &["-1"]
    );
}

#[test]
fn excflow_nested_28() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try try try raise Exception.Create('deep28'); except WriteLn('L328'); end; except WriteLn('L2'); end; except WriteLn('L1'); end; end."#
        ),
        &["L328"]
    );
}

#[test]
fn excflow_try_ok_29() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try WriteLn('try29'); except WriteLn('ex'); end; end."#
        ),
        &["try29"]
    );
}

#[test]
fn excflow_catch_raise_30() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try raise Exception.Create('boom30'); except WriteLn('got30'); end; end."#
        ),
        &["got30"]
    );
}

#[test]
fn excflow_finally_31() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try WriteLn('x31'); finally WriteLn('fin31'); end; end."#
        ),
        &["x31", "fin31"]
    );
}

#[test]
fn excflow_except_finally_32() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try try raise Exception.Create('e32'); except WriteLn('c32'); end; finally WriteLn('f32'); end; end."#
        ),
        &["c32", "f32"]
    );
}

#[test]
fn excflow_typed_33() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try raise ERangeError.Create('rng33'); except on E:ERangeError do WriteLn('range33'); on E:Exception do WriteLn('other'); end; end."#
        ),
        &["range33"]
    );
}

#[test]
fn excflow_div_zero_34() {
    assert_eq!(
        run_pascal(
            r#"program T; var v:Integer; begin v:=34; try v:=v div 0; except v:=-1; end; WriteLn(v); end."#
        ),
        &["-1"]
    );
}

#[test]
fn excflow_nested_35() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try try try raise Exception.Create('deep35'); except WriteLn('L335'); end; except WriteLn('L2'); end; except WriteLn('L1'); end; end."#
        ),
        &["L335"]
    );
}

#[test]
fn excflow_try_ok_36() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try WriteLn('try36'); except WriteLn('ex'); end; end."#
        ),
        &["try36"]
    );
}

#[test]
fn excflow_catch_raise_37() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try raise Exception.Create('boom37'); except WriteLn('got37'); end; end."#
        ),
        &["got37"]
    );
}

#[test]
fn excflow_finally_38() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try WriteLn('x38'); finally WriteLn('fin38'); end; end."#
        ),
        &["x38", "fin38"]
    );
}

#[test]
fn excflow_except_finally_39() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try try raise Exception.Create('e39'); except WriteLn('c39'); end; finally WriteLn('f39'); end; end."#
        ),
        &["c39", "f39"]
    );
}

#[test]
fn excflow_typed_40() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try raise ERangeError.Create('rng40'); except on E:ERangeError do WriteLn('range40'); on E:Exception do WriteLn('other'); end; end."#
        ),
        &["range40"]
    );
}

#[test]
fn excflow_div_zero_41() {
    assert_eq!(
        run_pascal(
            r#"program T; var v:Integer; begin v:=41; try v:=v div 0; except v:=-1; end; WriteLn(v); end."#
        ),
        &["-1"]
    );
}

#[test]
fn excflow_nested_42() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try try try raise Exception.Create('deep42'); except WriteLn('L342'); end; except WriteLn('L2'); end; except WriteLn('L1'); end; end."#
        ),
        &["L342"]
    );
}
