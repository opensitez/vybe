/// Program uses, var, const, and type section patterns.
use super::helpers::run_pascal;

#[test]
fn const_section_1() {
    assert_eq!(
        run_pascal(
            r#"program T; const K=1; begin WriteLn(K); end."#
        ),
        &["1"]
    );
}

#[test]
fn const_section_2() {
    assert_eq!(
        run_pascal(
            r#"program T; const K=2; begin WriteLn(K); end."#
        ),
        &["2"]
    );
}

#[test]
fn const_section_3() {
    assert_eq!(
        run_pascal(
            r#"program T; const K=3; begin WriteLn(K); end."#
        ),
        &["3"]
    );
}

#[test]
fn const_section_4() {
    assert_eq!(
        run_pascal(
            r#"program T; const K=4; begin WriteLn(K); end."#
        ),
        &["4"]
    );
}

#[test]
fn const_section_5() {
    assert_eq!(
        run_pascal(
            r#"program T; const K=5; begin WriteLn(K); end."#
        ),
        &["5"]
    );
}

#[test]
fn const_section_6() {
    assert_eq!(
        run_pascal(
            r#"program T; const K=6; begin WriteLn(K); end."#
        ),
        &["6"]
    );
}

#[test]
fn const_section_7() {
    assert_eq!(
        run_pascal(
            r#"program T; const K=7; begin WriteLn(K); end."#
        ),
        &["7"]
    );
}

#[test]
fn const_section_8() {
    assert_eq!(
        run_pascal(
            r#"program T; const K=8; begin WriteLn(K); end."#
        ),
        &["8"]
    );
}

#[test]
fn const_section_9() {
    assert_eq!(
        run_pascal(
            r#"program T; const K=9; begin WriteLn(K); end."#
        ),
        &["9"]
    );
}

#[test]
fn const_section_10() {
    assert_eq!(
        run_pascal(
            r#"program T; const K=10; begin WriteLn(K); end."#
        ),
        &["10"]
    );
}

#[test]
fn type_section_1() {
    assert_eq!(
        run_pascal(
            r#"program T; type TInt=Integer; var x:TInt; begin x:=1; WriteLn(x); end."#
        ),
        &["1"]
    );
}

#[test]
fn type_section_2() {
    assert_eq!(
        run_pascal(
            r#"program T; type TInt=Integer; var x:TInt; begin x:=2; WriteLn(x); end."#
        ),
        &["2"]
    );
}

#[test]
fn type_section_3() {
    assert_eq!(
        run_pascal(
            r#"program T; type TInt=Integer; var x:TInt; begin x:=3; WriteLn(x); end."#
        ),
        &["3"]
    );
}

#[test]
fn type_section_4() {
    assert_eq!(
        run_pascal(
            r#"program T; type TInt=Integer; var x:TInt; begin x:=4; WriteLn(x); end."#
        ),
        &["4"]
    );
}

#[test]
fn type_section_5() {
    assert_eq!(
        run_pascal(
            r#"program T; type TInt=Integer; var x:TInt; begin x:=5; WriteLn(x); end."#
        ),
        &["5"]
    );
}

#[test]
fn type_section_6() {
    assert_eq!(
        run_pascal(
            r#"program T; type TInt=Integer; var x:TInt; begin x:=6; WriteLn(x); end."#
        ),
        &["6"]
    );
}

#[test]
fn type_section_7() {
    assert_eq!(
        run_pascal(
            r#"program T; type TInt=Integer; var x:TInt; begin x:=7; WriteLn(x); end."#
        ),
        &["7"]
    );
}

#[test]
fn type_section_8() {
    assert_eq!(
        run_pascal(
            r#"program T; type TInt=Integer; var x:TInt; begin x:=8; WriteLn(x); end."#
        ),
        &["8"]
    );
}

#[test]
fn type_section_9() {
    assert_eq!(
        run_pascal(
            r#"program T; type TInt=Integer; var x:TInt; begin x:=9; WriteLn(x); end."#
        ),
        &["9"]
    );
}

#[test]
fn type_section_10() {
    assert_eq!(
        run_pascal(
            r#"program T; type TInt=Integer; var x:TInt; begin x:=10; WriteLn(x); end."#
        ),
        &["10"]
    );
}

#[test]
fn var_section_1() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Integer; begin a:=1; b:=2; WriteLn(a+b); end."#
        ),
        &["3"]
    );
}

#[test]
fn var_section_2() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Integer; begin a:=2; b:=3; WriteLn(a+b); end."#
        ),
        &["5"]
    );
}

#[test]
fn var_section_3() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Integer; begin a:=3; b:=4; WriteLn(a+b); end."#
        ),
        &["7"]
    );
}

#[test]
fn var_section_4() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Integer; begin a:=4; b:=5; WriteLn(a+b); end."#
        ),
        &["9"]
    );
}

#[test]
fn var_section_5() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Integer; begin a:=5; b:=6; WriteLn(a+b); end."#
        ),
        &["11"]
    );
}

#[test]
fn var_section_6() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Integer; begin a:=6; b:=7; WriteLn(a+b); end."#
        ),
        &["13"]
    );
}

#[test]
fn var_section_7() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Integer; begin a:=7; b:=8; WriteLn(a+b); end."#
        ),
        &["15"]
    );
}

#[test]
fn var_section_8() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Integer; begin a:=8; b:=9; WriteLn(a+b); end."#
        ),
        &["17"]
    );
}

#[test]
fn var_section_9() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Integer; begin a:=9; b:=10; WriteLn(a+b); end."#
        ),
        &["19"]
    );
}

#[test]
fn var_section_10() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Integer; begin a:=10; b:=11; WriteLn(a+b); end."#
        ),
        &["21"]
    );
}

#[test]
fn multi_section_1() {
    assert_eq!(
        run_pascal(
            r#"program T; const C=1; type TRec=record N:Integer; end; var r:TRec; begin r.N:=C*2; WriteLn(r.N); end."#
        ),
        &["2"]
    );
}

#[test]
fn multi_section_2() {
    assert_eq!(
        run_pascal(
            r#"program T; const C=2; type TRec=record N:Integer; end; var r:TRec; begin r.N:=C*2; WriteLn(r.N); end."#
        ),
        &["4"]
    );
}

#[test]
fn multi_section_3() {
    assert_eq!(
        run_pascal(
            r#"program T; const C=3; type TRec=record N:Integer; end; var r:TRec; begin r.N:=C*2; WriteLn(r.N); end."#
        ),
        &["6"]
    );
}

#[test]
fn multi_section_4() {
    assert_eq!(
        run_pascal(
            r#"program T; const C=4; type TRec=record N:Integer; end; var r:TRec; begin r.N:=C*2; WriteLn(r.N); end."#
        ),
        &["8"]
    );
}

#[test]
fn multi_section_5() {
    assert_eq!(
        run_pascal(
            r#"program T; const C=5; type TRec=record N:Integer; end; var r:TRec; begin r.N:=C*2; WriteLn(r.N); end."#
        ),
        &["10"]
    );
}

#[test]
fn multi_section_6() {
    assert_eq!(
        run_pascal(
            r#"program T; const C=6; type TRec=record N:Integer; end; var r:TRec; begin r.N:=C*2; WriteLn(r.N); end."#
        ),
        &["12"]
    );
}

#[test]
fn multi_section_7() {
    assert_eq!(
        run_pascal(
            r#"program T; const C=7; type TRec=record N:Integer; end; var r:TRec; begin r.N:=C*2; WriteLn(r.N); end."#
        ),
        &["14"]
    );
}

#[test]
fn multi_section_8() {
    assert_eq!(
        run_pascal(
            r#"program T; const C=8; type TRec=record N:Integer; end; var r:TRec; begin r.N:=C*2; WriteLn(r.N); end."#
        ),
        &["16"]
    );
}

#[test]
fn multi_section_9() {
    assert_eq!(
        run_pascal(
            r#"program T; const C=9; type TRec=record N:Integer; end; var r:TRec; begin r.N:=C*2; WriteLn(r.N); end."#
        ),
        &["18"]
    );
}

#[test]
fn multi_section_10() {
    assert_eq!(
        run_pascal(
            r#"program T; const C=10; type TRec=record N:Integer; end; var r:TRec; begin r.N:=C*2; WriteLn(r.N); end."#
        ),
        &["20"]
    );
}

#[test]
fn label_declaration_section() {
    assert_eq!(
        run_pascal(
            r#"program T; label finish; begin WriteLn('go'); goto finish; WriteLn('skip'); finish: WriteLn('done'); end."#
        ),
        &["go", "done"]
    );
}

#[test]
fn nested_local_procedure_in_main() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Outer; procedure Inner; begin WriteLn('inner'); end; begin Inner; end; begin Outer; end."#
        ),
        &["inner"]
    );
}

