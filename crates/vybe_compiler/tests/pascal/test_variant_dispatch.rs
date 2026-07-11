/// Variant type and variant record tag dispatch.
use super::helpers::run_pascal;

#[test]
fn vardisp_int_1() {
    assert_eq!(
        run_pascal(r#"program T; var v:Variant; begin v:=1; WriteLn(v); end."#),
        &["1"]
    );
}

#[test]
fn vardisp_str_2() {
    assert_eq!(
        run_pascal(r#"program T; var v:Variant; begin v:='txt2'; WriteLn(v); end."#),
        &["txt2"]
    );
}

#[test]
fn vardisp_rec_int_arm_3() {
    assert_eq!(
        run_pascal(
            r#"program T; type TV=record case K:Integer of 0:(I:Integer); 1:(S:string); end; var v:TV; begin v.K:=0; v.I:=3; WriteLn(v.I); end."#
        ),
        &["3"]
    );
}

#[test]
fn vardisp_rec_str_arm_4() {
    assert_eq!(
        run_pascal(
            r#"program T; type TV=record case K:Integer of 0:(I:Integer); 1:(S:string); end; var v:TV; begin v.K:=1; v.S:='arm4'; WriteLn(v.S); end."#
        ),
        &["arm4"]
    );
}

#[test]
fn vardisp_case_tag_5() {
    assert_eq!(
        run_pascal(
            r#"program T; type TV=record case K:Integer of 0:(A:Integer); 1:(B:Integer); 2:(C:Integer); end; procedure Show(const v:TV); begin case v.K of 0:WriteLn(v.A); 1:WriteLn(v.B); 2:WriteLn(v.C); end; end; var x:TV; begin x.K:=2; x.C:=35; Show(x); end."#
        ),
        &["35"]
    );
}

#[test]
fn vardisp_enum_tag_6() {
    assert_eq!(
        run_pascal(
            r#"program T; type TK=(One,Two,Three); type TV=record case Kind:TK of One:(N:Integer); Two:(T:string); Three:(C:Integer); end; var v:TV; begin v.Kind:=Three; v.C:=106; WriteLn(v.C); end."#
        ),
        &["106"]
    );
}

#[test]
fn vardisp_int_7() {
    assert_eq!(
        run_pascal(r#"program T; var v:Variant; begin v:=7; WriteLn(v); end."#),
        &["7"]
    );
}

#[test]
fn vardisp_str_8() {
    assert_eq!(
        run_pascal(r#"program T; var v:Variant; begin v:='txt8'; WriteLn(v); end."#),
        &["txt8"]
    );
}

#[test]
fn vardisp_rec_int_arm_9() {
    assert_eq!(
        run_pascal(
            r#"program T; type TV=record case K:Integer of 0:(I:Integer); 1:(S:string); end; var v:TV; begin v.K:=0; v.I:=9; WriteLn(v.I); end."#
        ),
        &["9"]
    );
}

#[test]
fn vardisp_rec_str_arm_10() {
    assert_eq!(
        run_pascal(
            r#"program T; type TV=record case K:Integer of 0:(I:Integer); 1:(S:string); end; var v:TV; begin v.K:=1; v.S:='arm10'; WriteLn(v.S); end."#
        ),
        &["arm10"]
    );
}

#[test]
fn vardisp_case_tag_11() {
    assert_eq!(
        run_pascal(
            r#"program T; type TV=record case K:Integer of 0:(A:Integer); 1:(B:Integer); 2:(C:Integer); end; procedure Show(const v:TV); begin case v.K of 0:WriteLn(v.A); 1:WriteLn(v.B); 2:WriteLn(v.C); end; end; var x:TV; begin x.K:=2; x.C:=77; Show(x); end."#
        ),
        &["77"]
    );
}

#[test]
fn vardisp_enum_tag_12() {
    assert_eq!(
        run_pascal(
            r#"program T; type TK=(One,Two,Three); type TV=record case Kind:TK of One:(N:Integer); Two:(T:string); Three:(C:Integer); end; var v:TV; begin v.Kind:=Three; v.C:=112; WriteLn(v.C); end."#
        ),
        &["112"]
    );
}

#[test]
fn vardisp_int_13() {
    assert_eq!(
        run_pascal(r#"program T; var v:Variant; begin v:=13; WriteLn(v); end."#),
        &["13"]
    );
}

#[test]
fn vardisp_str_14() {
    assert_eq!(
        run_pascal(r#"program T; var v:Variant; begin v:='txt14'; WriteLn(v); end."#),
        &["txt14"]
    );
}

#[test]
fn vardisp_rec_int_arm_15() {
    assert_eq!(
        run_pascal(
            r#"program T; type TV=record case K:Integer of 0:(I:Integer); 1:(S:string); end; var v:TV; begin v.K:=0; v.I:=15; WriteLn(v.I); end."#
        ),
        &["15"]
    );
}

#[test]
fn vardisp_rec_str_arm_16() {
    assert_eq!(
        run_pascal(
            r#"program T; type TV=record case K:Integer of 0:(I:Integer); 1:(S:string); end; var v:TV; begin v.K:=1; v.S:='arm16'; WriteLn(v.S); end."#
        ),
        &["arm16"]
    );
}

#[test]
fn vardisp_case_tag_17() {
    assert_eq!(
        run_pascal(
            r#"program T; type TV=record case K:Integer of 0:(A:Integer); 1:(B:Integer); 2:(C:Integer); end; procedure Show(const v:TV); begin case v.K of 0:WriteLn(v.A); 1:WriteLn(v.B); 2:WriteLn(v.C); end; end; var x:TV; begin x.K:=2; x.C:=119; Show(x); end."#
        ),
        &["119"]
    );
}

#[test]
fn vardisp_enum_tag_18() {
    assert_eq!(
        run_pascal(
            r#"program T; type TK=(One,Two,Three); type TV=record case Kind:TK of One:(N:Integer); Two:(T:string); Three:(C:Integer); end; var v:TV; begin v.Kind:=Three; v.C:=118; WriteLn(v.C); end."#
        ),
        &["118"]
    );
}

#[test]
fn vardisp_int_19() {
    assert_eq!(
        run_pascal(r#"program T; var v:Variant; begin v:=19; WriteLn(v); end."#),
        &["19"]
    );
}

#[test]
fn vardisp_str_20() {
    assert_eq!(
        run_pascal(r#"program T; var v:Variant; begin v:='txt20'; WriteLn(v); end."#),
        &["txt20"]
    );
}

#[test]
fn vardisp_rec_int_arm_21() {
    assert_eq!(
        run_pascal(
            r#"program T; type TV=record case K:Integer of 0:(I:Integer); 1:(S:string); end; var v:TV; begin v.K:=0; v.I:=21; WriteLn(v.I); end."#
        ),
        &["21"]
    );
}

#[test]
fn vardisp_rec_str_arm_22() {
    assert_eq!(
        run_pascal(
            r#"program T; type TV=record case K:Integer of 0:(I:Integer); 1:(S:string); end; var v:TV; begin v.K:=1; v.S:='arm22'; WriteLn(v.S); end."#
        ),
        &["arm22"]
    );
}

#[test]
fn vardisp_case_tag_23() {
    assert_eq!(
        run_pascal(
            r#"program T; type TV=record case K:Integer of 0:(A:Integer); 1:(B:Integer); 2:(C:Integer); end; procedure Show(const v:TV); begin case v.K of 0:WriteLn(v.A); 1:WriteLn(v.B); 2:WriteLn(v.C); end; end; var x:TV; begin x.K:=2; x.C:=161; Show(x); end."#
        ),
        &["161"]
    );
}

#[test]
fn vardisp_enum_tag_24() {
    assert_eq!(
        run_pascal(
            r#"program T; type TK=(One,Two,Three); type TV=record case Kind:TK of One:(N:Integer); Two:(T:string); Three:(C:Integer); end; var v:TV; begin v.Kind:=Three; v.C:=124; WriteLn(v.C); end."#
        ),
        &["124"]
    );
}

#[test]
fn vardisp_int_25() {
    assert_eq!(
        run_pascal(r#"program T; var v:Variant; begin v:=25; WriteLn(v); end."#),
        &["25"]
    );
}

#[test]
fn vardisp_str_26() {
    assert_eq!(
        run_pascal(r#"program T; var v:Variant; begin v:='txt26'; WriteLn(v); end."#),
        &["txt26"]
    );
}

#[test]
fn vardisp_rec_int_arm_27() {
    assert_eq!(
        run_pascal(
            r#"program T; type TV=record case K:Integer of 0:(I:Integer); 1:(S:string); end; var v:TV; begin v.K:=0; v.I:=27; WriteLn(v.I); end."#
        ),
        &["27"]
    );
}

#[test]
fn vardisp_rec_str_arm_28() {
    assert_eq!(
        run_pascal(
            r#"program T; type TV=record case K:Integer of 0:(I:Integer); 1:(S:string); end; var v:TV; begin v.K:=1; v.S:='arm28'; WriteLn(v.S); end."#
        ),
        &["arm28"]
    );
}

#[test]
fn vardisp_case_tag_29() {
    assert_eq!(
        run_pascal(
            r#"program T; type TV=record case K:Integer of 0:(A:Integer); 1:(B:Integer); 2:(C:Integer); end; procedure Show(const v:TV); begin case v.K of 0:WriteLn(v.A); 1:WriteLn(v.B); 2:WriteLn(v.C); end; end; var x:TV; begin x.K:=2; x.C:=203; Show(x); end."#
        ),
        &["203"]
    );
}

#[test]
fn vardisp_enum_tag_30() {
    assert_eq!(
        run_pascal(
            r#"program T; type TK=(One,Two,Three); type TV=record case Kind:TK of One:(N:Integer); Two:(T:string); Three:(C:Integer); end; var v:TV; begin v.Kind:=Three; v.C:=130; WriteLn(v.C); end."#
        ),
        &["130"]
    );
}

#[test]
fn vardisp_int_31() {
    assert_eq!(
        run_pascal(r#"program T; var v:Variant; begin v:=31; WriteLn(v); end."#),
        &["31"]
    );
}

#[test]
fn vardisp_str_32() {
    assert_eq!(
        run_pascal(r#"program T; var v:Variant; begin v:='txt32'; WriteLn(v); end."#),
        &["txt32"]
    );
}

#[test]
fn vardisp_rec_int_arm_33() {
    assert_eq!(
        run_pascal(
            r#"program T; type TV=record case K:Integer of 0:(I:Integer); 1:(S:string); end; var v:TV; begin v.K:=0; v.I:=33; WriteLn(v.I); end."#
        ),
        &["33"]
    );
}

#[test]
fn vardisp_rec_str_arm_34() {
    assert_eq!(
        run_pascal(
            r#"program T; type TV=record case K:Integer of 0:(I:Integer); 1:(S:string); end; var v:TV; begin v.K:=1; v.S:='arm34'; WriteLn(v.S); end."#
        ),
        &["arm34"]
    );
}

#[test]
fn vardisp_case_tag_35() {
    assert_eq!(
        run_pascal(
            r#"program T; type TV=record case K:Integer of 0:(A:Integer); 1:(B:Integer); 2:(C:Integer); end; procedure Show(const v:TV); begin case v.K of 0:WriteLn(v.A); 1:WriteLn(v.B); 2:WriteLn(v.C); end; end; var x:TV; begin x.K:=2; x.C:=245; Show(x); end."#
        ),
        &["245"]
    );
}

#[test]
fn vardisp_enum_tag_36() {
    assert_eq!(
        run_pascal(
            r#"program T; type TK=(One,Two,Three); type TV=record case Kind:TK of One:(N:Integer); Two:(T:string); Three:(C:Integer); end; var v:TV; begin v.Kind:=Three; v.C:=136; WriteLn(v.C); end."#
        ),
        &["136"]
    );
}

#[test]
fn vardisp_int_37() {
    assert_eq!(
        run_pascal(r#"program T; var v:Variant; begin v:=37; WriteLn(v); end."#),
        &["37"]
    );
}

#[test]
fn vardisp_str_38() {
    assert_eq!(
        run_pascal(r#"program T; var v:Variant; begin v:='txt38'; WriteLn(v); end."#),
        &["txt38"]
    );
}

#[test]
fn vardisp_rec_int_arm_39() {
    assert_eq!(
        run_pascal(
            r#"program T; type TV=record case K:Integer of 0:(I:Integer); 1:(S:string); end; var v:TV; begin v.K:=0; v.I:=39; WriteLn(v.I); end."#
        ),
        &["39"]
    );
}

#[test]
fn vardisp_rec_str_arm_40() {
    assert_eq!(
        run_pascal(
            r#"program T; type TV=record case K:Integer of 0:(I:Integer); 1:(S:string); end; var v:TV; begin v.K:=1; v.S:='arm40'; WriteLn(v.S); end."#
        ),
        &["arm40"]
    );
}

#[test]
fn vardisp_case_tag_41() {
    assert_eq!(
        run_pascal(
            r#"program T; type TV=record case K:Integer of 0:(A:Integer); 1:(B:Integer); 2:(C:Integer); end; procedure Show(const v:TV); begin case v.K of 0:WriteLn(v.A); 1:WriteLn(v.B); 2:WriteLn(v.C); end; end; var x:TV; begin x.K:=2; x.C:=287; Show(x); end."#
        ),
        &["287"]
    );
}

#[test]
fn vardisp_enum_tag_42() {
    assert_eq!(
        run_pascal(
            r#"program T; type TK=(One,Two,Three); type TV=record case Kind:TK of One:(N:Integer); Two:(T:string); Three:(C:Integer); end; var v:TV; begin v.Kind:=Three; v.C:=142; WriteLn(v.C); end."#
        ),
        &["142"]
    );
}
