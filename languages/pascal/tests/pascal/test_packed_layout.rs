/// Packed record byte alignment and layout scenarios.
use super::helpers::run_pascal;

#[test]
fn pklay_two_bytes_1() {
    assert_eq!(
        run_pascal(
            r#"program T; type TP=packed record A,B:Byte; end; var p:TP; begin p.A:=1; p.B:=2; WriteLn(p.A); WriteLn(p.B); end."#
        ),
        &["1", "2"]
    );
}

#[test]
fn pklay_bool_byte_2() {
    assert_eq!(
        run_pascal(
            r#"program T; type TP=packed record Flag:Boolean; Code:Byte; end; var p:TP; begin p.Flag:=true; p.Code:=2; if p.Flag then WriteLn(p.Code); end."#
        ),
        &["2"]
    );
}

#[test]
fn pklay_char_byte_3() {
    assert_eq!(
        run_pascal(
            r#"program T; type TP=packed record Ch:Char; B:Byte; end; var p:TP; begin p.Ch:='D'; p.B:=3; WriteLn(p.Ch); end."#
        ),
        &["D"]
    );
}

#[test]
fn pklay_nested_4() {
    assert_eq!(
        run_pascal(
            r#"program T; type TInner=packed record B:Byte; end; type TOuter=packed record Inner:TInner; Tag:Byte; end; var o:TOuter; begin o.Inner.B:=4; o.Tag:=9; WriteLn(o.Inner.B); WriteLn(o.Tag); end."#
        ),
        &["4", "9"]
    );
}

#[test]
fn pklay_case_ver_5() {
    assert_eq!(
        run_pascal(
            r#"program T; type TP=packed record Ver:Byte; end; var p:TP; begin p.Ver:=3; case p.Ver of 1:WriteLn('v1_5'); 2:WriteLn('v2_5'); 3:WriteLn('v3_5'); else WriteLn('?'); end; end."#
        ),
        &["v3_5"]
    );
}

#[test]
fn pklay_assign_sum_6() {
    assert_eq!(
        run_pascal(
            r#"program T; type TP=packed record Lo,Hi:Byte; end; var a,b:TP; begin a.Lo:=6; a.Hi:=8; b:=a; WriteLn(b.Lo+b.Hi); end."#
        ),
        &["14"]
    );
}

#[test]
fn pklay_two_bytes_7() {
    assert_eq!(
        run_pascal(
            r#"program T; type TP=packed record A,B:Byte; end; var p:TP; begin p.A:=7; p.B:=8; WriteLn(p.A); WriteLn(p.B); end."#
        ),
        &["7", "8"]
    );
}

#[test]
fn pklay_bool_byte_8() {
    assert_eq!(
        run_pascal(
            r#"program T; type TP=packed record Flag:Boolean; Code:Byte; end; var p:TP; begin p.Flag:=true; p.Code:=8; if p.Flag then WriteLn(p.Code); end."#
        ),
        &["8"]
    );
}

#[test]
fn pklay_char_byte_9() {
    assert_eq!(
        run_pascal(
            r#"program T; type TP=packed record Ch:Char; B:Byte; end; var p:TP; begin p.Ch:='J'; p.B:=9; WriteLn(p.Ch); end."#
        ),
        &["J"]
    );
}

#[test]
fn pklay_nested_10() {
    assert_eq!(
        run_pascal(
            r#"program T; type TInner=packed record B:Byte; end; type TOuter=packed record Inner:TInner; Tag:Byte; end; var o:TOuter; begin o.Inner.B:=10; o.Tag:=15; WriteLn(o.Inner.B); WriteLn(o.Tag); end."#
        ),
        &["10", "15"]
    );
}

#[test]
fn pklay_case_ver_11() {
    assert_eq!(
        run_pascal(
            r#"program T; type TP=packed record Ver:Byte; end; var p:TP; begin p.Ver:=3; case p.Ver of 1:WriteLn('v1_11'); 2:WriteLn('v2_11'); 3:WriteLn('v3_11'); else WriteLn('?'); end; end."#
        ),
        &["v3_11"]
    );
}

#[test]
fn pklay_assign_sum_12() {
    assert_eq!(
        run_pascal(
            r#"program T; type TP=packed record Lo,Hi:Byte; end; var a,b:TP; begin a.Lo:=12; a.Hi:=14; b:=a; WriteLn(b.Lo+b.Hi); end."#
        ),
        &["26"]
    );
}

#[test]
fn pklay_two_bytes_13() {
    assert_eq!(
        run_pascal(
            r#"program T; type TP=packed record A,B:Byte; end; var p:TP; begin p.A:=13; p.B:=14; WriteLn(p.A); WriteLn(p.B); end."#
        ),
        &["13", "14"]
    );
}

#[test]
fn pklay_bool_byte_14() {
    assert_eq!(
        run_pascal(
            r#"program T; type TP=packed record Flag:Boolean; Code:Byte; end; var p:TP; begin p.Flag:=true; p.Code:=14; if p.Flag then WriteLn(p.Code); end."#
        ),
        &["14"]
    );
}

#[test]
fn pklay_char_byte_15() {
    assert_eq!(
        run_pascal(
            r#"program T; type TP=packed record Ch:Char; B:Byte; end; var p:TP; begin p.Ch:='P'; p.B:=15; WriteLn(p.Ch); end."#
        ),
        &["P"]
    );
}

#[test]
fn pklay_nested_16() {
    assert_eq!(
        run_pascal(
            r#"program T; type TInner=packed record B:Byte; end; type TOuter=packed record Inner:TInner; Tag:Byte; end; var o:TOuter; begin o.Inner.B:=16; o.Tag:=21; WriteLn(o.Inner.B); WriteLn(o.Tag); end."#
        ),
        &["16", "21"]
    );
}

#[test]
fn pklay_case_ver_17() {
    assert_eq!(
        run_pascal(
            r#"program T; type TP=packed record Ver:Byte; end; var p:TP; begin p.Ver:=3; case p.Ver of 1:WriteLn('v1_17'); 2:WriteLn('v2_17'); 3:WriteLn('v3_17'); else WriteLn('?'); end; end."#
        ),
        &["v3_17"]
    );
}

#[test]
fn pklay_assign_sum_18() {
    assert_eq!(
        run_pascal(
            r#"program T; type TP=packed record Lo,Hi:Byte; end; var a,b:TP; begin a.Lo:=18; a.Hi:=20; b:=a; WriteLn(b.Lo+b.Hi); end."#
        ),
        &["38"]
    );
}

#[test]
fn pklay_two_bytes_19() {
    assert_eq!(
        run_pascal(
            r#"program T; type TP=packed record A,B:Byte; end; var p:TP; begin p.A:=19; p.B:=20; WriteLn(p.A); WriteLn(p.B); end."#
        ),
        &["19", "20"]
    );
}

#[test]
fn pklay_bool_byte_20() {
    assert_eq!(
        run_pascal(
            r#"program T; type TP=packed record Flag:Boolean; Code:Byte; end; var p:TP; begin p.Flag:=true; p.Code:=20; if p.Flag then WriteLn(p.Code); end."#
        ),
        &["20"]
    );
}

#[test]
fn pklay_char_byte_21() {
    assert_eq!(
        run_pascal(
            r#"program T; type TP=packed record Ch:Char; B:Byte; end; var p:TP; begin p.Ch:='V'; p.B:=21; WriteLn(p.Ch); end."#
        ),
        &["V"]
    );
}

#[test]
fn pklay_nested_22() {
    assert_eq!(
        run_pascal(
            r#"program T; type TInner=packed record B:Byte; end; type TOuter=packed record Inner:TInner; Tag:Byte; end; var o:TOuter; begin o.Inner.B:=22; o.Tag:=27; WriteLn(o.Inner.B); WriteLn(o.Tag); end."#
        ),
        &["22", "27"]
    );
}

#[test]
fn pklay_case_ver_23() {
    assert_eq!(
        run_pascal(
            r#"program T; type TP=packed record Ver:Byte; end; var p:TP; begin p.Ver:=3; case p.Ver of 1:WriteLn('v1_23'); 2:WriteLn('v2_23'); 3:WriteLn('v3_23'); else WriteLn('?'); end; end."#
        ),
        &["v3_23"]
    );
}

#[test]
fn pklay_assign_sum_24() {
    assert_eq!(
        run_pascal(
            r#"program T; type TP=packed record Lo,Hi:Byte; end; var a,b:TP; begin a.Lo:=24; a.Hi:=26; b:=a; WriteLn(b.Lo+b.Hi); end."#
        ),
        &["50"]
    );
}

#[test]
fn pklay_two_bytes_25() {
    assert_eq!(
        run_pascal(
            r#"program T; type TP=packed record A,B:Byte; end; var p:TP; begin p.A:=25; p.B:=26; WriteLn(p.A); WriteLn(p.B); end."#
        ),
        &["25", "26"]
    );
}

#[test]
fn pklay_bool_byte_26() {
    assert_eq!(
        run_pascal(
            r#"program T; type TP=packed record Flag:Boolean; Code:Byte; end; var p:TP; begin p.Flag:=true; p.Code:=26; if p.Flag then WriteLn(p.Code); end."#
        ),
        &["26"]
    );
}

#[test]
fn pklay_char_byte_27() {
    assert_eq!(
        run_pascal(
            r#"program T; type TP=packed record Ch:Char; B:Byte; end; var p:TP; begin p.Ch:='B'; p.B:=27; WriteLn(p.Ch); end."#
        ),
        &["B"]
    );
}

#[test]
fn pklay_nested_28() {
    assert_eq!(
        run_pascal(
            r#"program T; type TInner=packed record B:Byte; end; type TOuter=packed record Inner:TInner; Tag:Byte; end; var o:TOuter; begin o.Inner.B:=28; o.Tag:=33; WriteLn(o.Inner.B); WriteLn(o.Tag); end."#
        ),
        &["28", "33"]
    );
}

#[test]
fn pklay_case_ver_29() {
    assert_eq!(
        run_pascal(
            r#"program T; type TP=packed record Ver:Byte; end; var p:TP; begin p.Ver:=3; case p.Ver of 1:WriteLn('v1_29'); 2:WriteLn('v2_29'); 3:WriteLn('v3_29'); else WriteLn('?'); end; end."#
        ),
        &["v3_29"]
    );
}

#[test]
fn pklay_assign_sum_30() {
    assert_eq!(
        run_pascal(
            r#"program T; type TP=packed record Lo,Hi:Byte; end; var a,b:TP; begin a.Lo:=30; a.Hi:=32; b:=a; WriteLn(b.Lo+b.Hi); end."#
        ),
        &["62"]
    );
}

#[test]
fn pklay_two_bytes_31() {
    assert_eq!(
        run_pascal(
            r#"program T; type TP=packed record A,B:Byte; end; var p:TP; begin p.A:=31; p.B:=32; WriteLn(p.A); WriteLn(p.B); end."#
        ),
        &["31", "32"]
    );
}

#[test]
fn pklay_bool_byte_32() {
    assert_eq!(
        run_pascal(
            r#"program T; type TP=packed record Flag:Boolean; Code:Byte; end; var p:TP; begin p.Flag:=true; p.Code:=32; if p.Flag then WriteLn(p.Code); end."#
        ),
        &["32"]
    );
}

#[test]
fn pklay_char_byte_33() {
    assert_eq!(
        run_pascal(
            r#"program T; type TP=packed record Ch:Char; B:Byte; end; var p:TP; begin p.Ch:='H'; p.B:=33; WriteLn(p.Ch); end."#
        ),
        &["H"]
    );
}

#[test]
fn pklay_nested_34() {
    assert_eq!(
        run_pascal(
            r#"program T; type TInner=packed record B:Byte; end; type TOuter=packed record Inner:TInner; Tag:Byte; end; var o:TOuter; begin o.Inner.B:=34; o.Tag:=39; WriteLn(o.Inner.B); WriteLn(o.Tag); end."#
        ),
        &["34", "39"]
    );
}

#[test]
fn pklay_case_ver_35() {
    assert_eq!(
        run_pascal(
            r#"program T; type TP=packed record Ver:Byte; end; var p:TP; begin p.Ver:=3; case p.Ver of 1:WriteLn('v1_35'); 2:WriteLn('v2_35'); 3:WriteLn('v3_35'); else WriteLn('?'); end; end."#
        ),
        &["v3_35"]
    );
}

#[test]
fn pklay_assign_sum_36() {
    assert_eq!(
        run_pascal(
            r#"program T; type TP=packed record Lo,Hi:Byte; end; var a,b:TP; begin a.Lo:=36; a.Hi:=38; b:=a; WriteLn(b.Lo+b.Hi); end."#
        ),
        &["74"]
    );
}

#[test]
fn pklay_two_bytes_37() {
    assert_eq!(
        run_pascal(
            r#"program T; type TP=packed record A,B:Byte; end; var p:TP; begin p.A:=37; p.B:=38; WriteLn(p.A); WriteLn(p.B); end."#
        ),
        &["37", "38"]
    );
}

#[test]
fn pklay_bool_byte_38() {
    assert_eq!(
        run_pascal(
            r#"program T; type TP=packed record Flag:Boolean; Code:Byte; end; var p:TP; begin p.Flag:=true; p.Code:=38; if p.Flag then WriteLn(p.Code); end."#
        ),
        &["38"]
    );
}

#[test]
fn pklay_char_byte_39() {
    assert_eq!(
        run_pascal(
            r#"program T; type TP=packed record Ch:Char; B:Byte; end; var p:TP; begin p.Ch:='N'; p.B:=39; WriteLn(p.Ch); end."#
        ),
        &["N"]
    );
}

#[test]
fn pklay_nested_40() {
    assert_eq!(
        run_pascal(
            r#"program T; type TInner=packed record B:Byte; end; type TOuter=packed record Inner:TInner; Tag:Byte; end; var o:TOuter; begin o.Inner.B:=40; o.Tag:=45; WriteLn(o.Inner.B); WriteLn(o.Tag); end."#
        ),
        &["40", "45"]
    );
}

#[test]
fn pklay_case_ver_41() {
    assert_eq!(
        run_pascal(
            r#"program T; type TP=packed record Ver:Byte; end; var p:TP; begin p.Ver:=3; case p.Ver of 1:WriteLn('v1_41'); 2:WriteLn('v2_41'); 3:WriteLn('v3_41'); else WriteLn('?'); end; end."#
        ),
        &["v3_41"]
    );
}

#[test]
fn pklay_assign_sum_42() {
    assert_eq!(
        run_pascal(
            r#"program T; type TP=packed record Lo,Hi:Byte; end; var a,b:TP; begin a.Lo:=42; a.Hi:=44; b:=a; WriteLn(b.Lo+b.Hi); end."#
        ),
        &["86"]
    );
}
