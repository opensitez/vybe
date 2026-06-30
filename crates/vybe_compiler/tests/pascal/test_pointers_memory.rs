/// GetMem/FreeMem and New/Dispose heap patterns.
use super::helpers::run_pascal;

#[test]
fn ptrmem_new_int_1() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:^Integer; begin New(p); p^:=1; WriteLn(p^); Dispose(p); end."#
        ),
        &["1"]
    );
}

#[test]
fn ptrmem_getmem_2() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:Pointer; begin GetMem(p,SizeOf(Integer)); PInteger(p)^:=4; WriteLn(PInteger(p)^); FreeMem(p); end."#
        ),
        &["4"]
    );
}

#[test]
fn ptrmem_new_record_3() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; end; var p:^TR; begin New(p); p^.V:=13; WriteLn(p^.V); Dispose(p); end."#
        ),
        &["13"]
    );
}

#[test]
fn ptrmem_two_ptrs_4() {
    assert_eq!(
        run_pascal(
            r#"program T; var p,q:^Integer; begin New(p); New(q); p^:=4; q^:=5; WriteLn(p^+q^); Dispose(p); Dispose(q); end."#
        ),
        &["9"]
    );
}

#[test]
fn ptrmem_new_string_5() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:^String; begin New(p); p^:='s5'; WriteLn(p^); Dispose(p); end."#
        ),
        &["s5"]
    );
}

#[test]
fn ptrmem_renew_6() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:^Integer; begin New(p); p^:=6; Dispose(p); New(p); p^:=18; WriteLn(p^); Dispose(p); end."#
        ),
        &["18"]
    );
}

#[test]
fn ptrmem_new_int_7() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:^Integer; begin New(p); p^:=7; WriteLn(p^); Dispose(p); end."#
        ),
        &["7"]
    );
}

#[test]
fn ptrmem_getmem_8() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:Pointer; begin GetMem(p,SizeOf(Integer)); PInteger(p)^:=16; WriteLn(PInteger(p)^); FreeMem(p); end."#
        ),
        &["16"]
    );
}

#[test]
fn ptrmem_new_record_9() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; end; var p:^TR; begin New(p); p^.V:=19; WriteLn(p^.V); Dispose(p); end."#
        ),
        &["19"]
    );
}

#[test]
fn ptrmem_two_ptrs_10() {
    assert_eq!(
        run_pascal(
            r#"program T; var p,q:^Integer; begin New(p); New(q); p^:=10; q^:=11; WriteLn(p^+q^); Dispose(p); Dispose(q); end."#
        ),
        &["21"]
    );
}

#[test]
fn ptrmem_new_string_11() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:^String; begin New(p); p^:='s11'; WriteLn(p^); Dispose(p); end."#
        ),
        &["s11"]
    );
}

#[test]
fn ptrmem_renew_12() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:^Integer; begin New(p); p^:=12; Dispose(p); New(p); p^:=36; WriteLn(p^); Dispose(p); end."#
        ),
        &["36"]
    );
}

#[test]
fn ptrmem_new_int_13() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:^Integer; begin New(p); p^:=13; WriteLn(p^); Dispose(p); end."#
        ),
        &["13"]
    );
}

#[test]
fn ptrmem_getmem_14() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:Pointer; begin GetMem(p,SizeOf(Integer)); PInteger(p)^:=28; WriteLn(PInteger(p)^); FreeMem(p); end."#
        ),
        &["28"]
    );
}

#[test]
fn ptrmem_new_record_15() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; end; var p:^TR; begin New(p); p^.V:=25; WriteLn(p^.V); Dispose(p); end."#
        ),
        &["25"]
    );
}

#[test]
fn ptrmem_two_ptrs_16() {
    assert_eq!(
        run_pascal(
            r#"program T; var p,q:^Integer; begin New(p); New(q); p^:=16; q^:=17; WriteLn(p^+q^); Dispose(p); Dispose(q); end."#
        ),
        &["33"]
    );
}

#[test]
fn ptrmem_new_string_17() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:^String; begin New(p); p^:='s17'; WriteLn(p^); Dispose(p); end."#
        ),
        &["s17"]
    );
}

#[test]
fn ptrmem_renew_18() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:^Integer; begin New(p); p^:=18; Dispose(p); New(p); p^:=54; WriteLn(p^); Dispose(p); end."#
        ),
        &["54"]
    );
}

#[test]
fn ptrmem_new_int_19() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:^Integer; begin New(p); p^:=19; WriteLn(p^); Dispose(p); end."#
        ),
        &["19"]
    );
}

#[test]
fn ptrmem_getmem_20() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:Pointer; begin GetMem(p,SizeOf(Integer)); PInteger(p)^:=40; WriteLn(PInteger(p)^); FreeMem(p); end."#
        ),
        &["40"]
    );
}

#[test]
fn ptrmem_new_record_21() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; end; var p:^TR; begin New(p); p^.V:=31; WriteLn(p^.V); Dispose(p); end."#
        ),
        &["31"]
    );
}

#[test]
fn ptrmem_two_ptrs_22() {
    assert_eq!(
        run_pascal(
            r#"program T; var p,q:^Integer; begin New(p); New(q); p^:=22; q^:=23; WriteLn(p^+q^); Dispose(p); Dispose(q); end."#
        ),
        &["45"]
    );
}

#[test]
fn ptrmem_new_string_23() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:^String; begin New(p); p^:='s23'; WriteLn(p^); Dispose(p); end."#
        ),
        &["s23"]
    );
}

#[test]
fn ptrmem_renew_24() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:^Integer; begin New(p); p^:=24; Dispose(p); New(p); p^:=72; WriteLn(p^); Dispose(p); end."#
        ),
        &["72"]
    );
}

#[test]
fn ptrmem_new_int_25() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:^Integer; begin New(p); p^:=25; WriteLn(p^); Dispose(p); end."#
        ),
        &["25"]
    );
}

#[test]
fn ptrmem_getmem_26() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:Pointer; begin GetMem(p,SizeOf(Integer)); PInteger(p)^:=52; WriteLn(PInteger(p)^); FreeMem(p); end."#
        ),
        &["52"]
    );
}

#[test]
fn ptrmem_new_record_27() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; end; var p:^TR; begin New(p); p^.V:=37; WriteLn(p^.V); Dispose(p); end."#
        ),
        &["37"]
    );
}

#[test]
fn ptrmem_two_ptrs_28() {
    assert_eq!(
        run_pascal(
            r#"program T; var p,q:^Integer; begin New(p); New(q); p^:=28; q^:=29; WriteLn(p^+q^); Dispose(p); Dispose(q); end."#
        ),
        &["57"]
    );
}

#[test]
fn ptrmem_new_string_29() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:^String; begin New(p); p^:='s29'; WriteLn(p^); Dispose(p); end."#
        ),
        &["s29"]
    );
}

#[test]
fn ptrmem_renew_30() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:^Integer; begin New(p); p^:=30; Dispose(p); New(p); p^:=90; WriteLn(p^); Dispose(p); end."#
        ),
        &["90"]
    );
}

#[test]
fn ptrmem_new_int_31() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:^Integer; begin New(p); p^:=31; WriteLn(p^); Dispose(p); end."#
        ),
        &["31"]
    );
}

#[test]
fn ptrmem_getmem_32() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:Pointer; begin GetMem(p,SizeOf(Integer)); PInteger(p)^:=64; WriteLn(PInteger(p)^); FreeMem(p); end."#
        ),
        &["64"]
    );
}

#[test]
fn ptrmem_new_record_33() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; end; var p:^TR; begin New(p); p^.V:=43; WriteLn(p^.V); Dispose(p); end."#
        ),
        &["43"]
    );
}

#[test]
fn ptrmem_two_ptrs_34() {
    assert_eq!(
        run_pascal(
            r#"program T; var p,q:^Integer; begin New(p); New(q); p^:=34; q^:=35; WriteLn(p^+q^); Dispose(p); Dispose(q); end."#
        ),
        &["69"]
    );
}

#[test]
fn ptrmem_new_string_35() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:^String; begin New(p); p^:='s35'; WriteLn(p^); Dispose(p); end."#
        ),
        &["s35"]
    );
}

#[test]
fn ptrmem_renew_36() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:^Integer; begin New(p); p^:=36; Dispose(p); New(p); p^:=108; WriteLn(p^); Dispose(p); end."#
        ),
        &["108"]
    );
}

#[test]
fn ptrmem_new_int_37() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:^Integer; begin New(p); p^:=37; WriteLn(p^); Dispose(p); end."#
        ),
        &["37"]
    );
}

#[test]
fn ptrmem_getmem_38() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:Pointer; begin GetMem(p,SizeOf(Integer)); PInteger(p)^:=76; WriteLn(PInteger(p)^); FreeMem(p); end."#
        ),
        &["76"]
    );
}

#[test]
fn ptrmem_new_record_39() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; end; var p:^TR; begin New(p); p^.V:=49; WriteLn(p^.V); Dispose(p); end."#
        ),
        &["49"]
    );
}

#[test]
fn ptrmem_two_ptrs_40() {
    assert_eq!(
        run_pascal(
            r#"program T; var p,q:^Integer; begin New(p); New(q); p^:=40; q^:=41; WriteLn(p^+q^); Dispose(p); Dispose(q); end."#
        ),
        &["81"]
    );
}

#[test]
fn ptrmem_new_string_41() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:^String; begin New(p); p^:='s41'; WriteLn(p^); Dispose(p); end."#
        ),
        &["s41"]
    );
}

#[test]
fn ptrmem_renew_42() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:^Integer; begin New(p); p^:=42; Dispose(p); New(p); p^:=126; WriteLn(p^); Dispose(p); end."#
        ),
        &["126"]
    );
}
