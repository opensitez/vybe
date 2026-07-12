/// Private, protected, public, and strict visibility on classes.
use super::helpers::run_pascal;

#[test]
fn public_field_1() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public F:Integer; end; var o:TB; begin o:=TB.Create; o.F:=1; WriteLn(o.F); o.Free; end."#
        ),
        &["1"]
    );
}

#[test]
fn public_field_2() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public F:Integer; end; var o:TB; begin o:=TB.Create; o.F:=2; WriteLn(o.F); o.Free; end."#
        ),
        &["2"]
    );
}

#[test]
fn public_field_3() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public F:Integer; end; var o:TB; begin o:=TB.Create; o.F:=3; WriteLn(o.F); o.Free; end."#
        ),
        &["3"]
    );
}

#[test]
fn public_field_4() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public F:Integer; end; var o:TB; begin o:=TB.Create; o.F:=4; WriteLn(o.F); o.Free; end."#
        ),
        &["4"]
    );
}

#[test]
fn public_field_5() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public F:Integer; end; var o:TB; begin o:=TB.Create; o.F:=5; WriteLn(o.F); o.Free; end."#
        ),
        &["5"]
    );
}

#[test]
fn public_field_6() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public F:Integer; end; var o:TB; begin o:=TB.Create; o.F:=6; WriteLn(o.F); o.Free; end."#
        ),
        &["6"]
    );
}

#[test]
fn public_field_7() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public F:Integer; end; var o:TB; begin o:=TB.Create; o.F:=7; WriteLn(o.F); o.Free; end."#
        ),
        &["7"]
    );
}

#[test]
fn public_field_8() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public F:Integer; end; var o:TB; begin o:=TB.Create; o.F:=8; WriteLn(o.F); o.Free; end."#
        ),
        &["8"]
    );
}

#[test]
fn public_field_9() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public F:Integer; end; var o:TB; begin o:=TB.Create; o.F:=9; WriteLn(o.F); o.Free; end."#
        ),
        &["9"]
    );
}

#[test]
fn public_field_10() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public F:Integer; end; var o:TB; begin o:=TB.Create; o.F:=10; WriteLn(o.F); o.Free; end."#
        ),
        &["10"]
    );
}

#[test]
fn public_field_11() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public F:Integer; end; var o:TB; begin o:=TB.Create; o.F:=11; WriteLn(o.F); o.Free; end."#
        ),
        &["11"]
    );
}

#[test]
fn public_field_12() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public F:Integer; end; var o:TB; begin o:=TB.Create; o.F:=12; WriteLn(o.F); o.Free; end."#
        ),
        &["12"]
    );
}

#[test]
fn public_field_13() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public F:Integer; end; var o:TB; begin o:=TB.Create; o.F:=13; WriteLn(o.F); o.Free; end."#
        ),
        &["13"]
    );
}

#[test]
fn public_field_14() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public F:Integer; end; var o:TB; begin o:=TB.Create; o.F:=14; WriteLn(o.F); o.Free; end."#
        ),
        &["14"]
    );
}

#[test]
fn strict_private_ctor_1() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class strict private F:Integer; public constructor Create(v:Integer); function Get:Integer; end; constructor TB.Create(v:Integer); begin F:=v; end; function TB.Get:Integer; begin Result:=F; end; var o:TB; begin o:=TB.Create(1); WriteLn(o.Get); o.Free; end."#
        ),
        &["1"]
    );
}

#[test]
fn strict_private_ctor_2() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class strict private F:Integer; public constructor Create(v:Integer); function Get:Integer; end; constructor TB.Create(v:Integer); begin F:=v; end; function TB.Get:Integer; begin Result:=F; end; var o:TB; begin o:=TB.Create(2); WriteLn(o.Get); o.Free; end."#
        ),
        &["2"]
    );
}

#[test]
fn strict_private_ctor_3() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class strict private F:Integer; public constructor Create(v:Integer); function Get:Integer; end; constructor TB.Create(v:Integer); begin F:=v; end; function TB.Get:Integer; begin Result:=F; end; var o:TB; begin o:=TB.Create(3); WriteLn(o.Get); o.Free; end."#
        ),
        &["3"]
    );
}

#[test]
fn strict_private_ctor_4() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class strict private F:Integer; public constructor Create(v:Integer); function Get:Integer; end; constructor TB.Create(v:Integer); begin F:=v; end; function TB.Get:Integer; begin Result:=F; end; var o:TB; begin o:=TB.Create(4); WriteLn(o.Get); o.Free; end."#
        ),
        &["4"]
    );
}

#[test]
fn strict_private_ctor_5() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class strict private F:Integer; public constructor Create(v:Integer); function Get:Integer; end; constructor TB.Create(v:Integer); begin F:=v; end; function TB.Get:Integer; begin Result:=F; end; var o:TB; begin o:=TB.Create(5); WriteLn(o.Get); o.Free; end."#
        ),
        &["5"]
    );
}

#[test]
fn strict_private_ctor_6() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class strict private F:Integer; public constructor Create(v:Integer); function Get:Integer; end; constructor TB.Create(v:Integer); begin F:=v; end; function TB.Get:Integer; begin Result:=F; end; var o:TB; begin o:=TB.Create(6); WriteLn(o.Get); o.Free; end."#
        ),
        &["6"]
    );
}

#[test]
fn strict_private_ctor_7() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class strict private F:Integer; public constructor Create(v:Integer); function Get:Integer; end; constructor TB.Create(v:Integer); begin F:=v; end; function TB.Get:Integer; begin Result:=F; end; var o:TB; begin o:=TB.Create(7); WriteLn(o.Get); o.Free; end."#
        ),
        &["7"]
    );
}

#[test]
fn strict_private_ctor_8() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class strict private F:Integer; public constructor Create(v:Integer); function Get:Integer; end; constructor TB.Create(v:Integer); begin F:=v; end; function TB.Get:Integer; begin Result:=F; end; var o:TB; begin o:=TB.Create(8); WriteLn(o.Get); o.Free; end."#
        ),
        &["8"]
    );
}

#[test]
fn strict_private_ctor_9() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class strict private F:Integer; public constructor Create(v:Integer); function Get:Integer; end; constructor TB.Create(v:Integer); begin F:=v; end; function TB.Get:Integer; begin Result:=F; end; var o:TB; begin o:=TB.Create(9); WriteLn(o.Get); o.Free; end."#
        ),
        &["9"]
    );
}

#[test]
fn strict_private_ctor_10() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class strict private F:Integer; public constructor Create(v:Integer); function Get:Integer; end; constructor TB.Create(v:Integer); begin F:=v; end; function TB.Get:Integer; begin Result:=F; end; var o:TB; begin o:=TB.Create(10); WriteLn(o.Get); o.Free; end."#
        ),
        &["10"]
    );
}

#[test]
fn strict_private_ctor_11() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class strict private F:Integer; public constructor Create(v:Integer); function Get:Integer; end; constructor TB.Create(v:Integer); begin F:=v; end; function TB.Get:Integer; begin Result:=F; end; var o:TB; begin o:=TB.Create(11); WriteLn(o.Get); o.Free; end."#
        ),
        &["11"]
    );
}

#[test]
fn strict_private_ctor_12() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class strict private F:Integer; public constructor Create(v:Integer); function Get:Integer; end; constructor TB.Create(v:Integer); begin F:=v; end; function TB.Get:Integer; begin Result:=F; end; var o:TB; begin o:=TB.Create(12); WriteLn(o.Get); o.Free; end."#
        ),
        &["12"]
    );
}

#[test]
fn strict_private_ctor_13() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class strict private F:Integer; public constructor Create(v:Integer); function Get:Integer; end; constructor TB.Create(v:Integer); begin F:=v; end; function TB.Get:Integer; begin Result:=F; end; var o:TB; begin o:=TB.Create(13); WriteLn(o.Get); o.Free; end."#
        ),
        &["13"]
    );
}

#[test]
fn strict_private_ctor_14() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class strict private F:Integer; public constructor Create(v:Integer); function Get:Integer; end; constructor TB.Create(v:Integer); begin F:=v; end; function TB.Get:Integer; begin Result:=F; end; var o:TB; begin o:=TB.Create(14); WriteLn(o.Get); o.Free; end."#
        ),
        &["14"]
    );
}

#[test]
fn virtual_override_tag_1() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public function Tag:Integer; virtual; end; TC=class(TB) function Tag:Integer; override; end; function TB.Tag:Integer; begin Result:=1; end; function TC.Tag:Integer; begin Result:=2; end; var o:TB; begin o:=TC.Create; WriteLn(o.Tag); o.Free; end."#
        ),
        &["2"]
    );
}

#[test]
fn virtual_override_tag_2() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public function Tag:Integer; virtual; end; TC=class(TB) function Tag:Integer; override; end; function TB.Tag:Integer; begin Result:=2; end; function TC.Tag:Integer; begin Result:=4; end; var o:TB; begin o:=TC.Create; WriteLn(o.Tag); o.Free; end."#
        ),
        &["4"]
    );
}

#[test]
fn virtual_override_tag_3() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public function Tag:Integer; virtual; end; TC=class(TB) function Tag:Integer; override; end; function TB.Tag:Integer; begin Result:=3; end; function TC.Tag:Integer; begin Result:=6; end; var o:TB; begin o:=TC.Create; WriteLn(o.Tag); o.Free; end."#
        ),
        &["6"]
    );
}

#[test]
fn virtual_override_tag_4() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public function Tag:Integer; virtual; end; TC=class(TB) function Tag:Integer; override; end; function TB.Tag:Integer; begin Result:=4; end; function TC.Tag:Integer; begin Result:=8; end; var o:TB; begin o:=TC.Create; WriteLn(o.Tag); o.Free; end."#
        ),
        &["8"]
    );
}

#[test]
fn virtual_override_tag_5() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public function Tag:Integer; virtual; end; TC=class(TB) function Tag:Integer; override; end; function TB.Tag:Integer; begin Result:=5; end; function TC.Tag:Integer; begin Result:=10; end; var o:TB; begin o:=TC.Create; WriteLn(o.Tag); o.Free; end."#
        ),
        &["10"]
    );
}

#[test]
fn virtual_override_tag_6() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public function Tag:Integer; virtual; end; TC=class(TB) function Tag:Integer; override; end; function TB.Tag:Integer; begin Result:=6; end; function TC.Tag:Integer; begin Result:=12; end; var o:TB; begin o:=TC.Create; WriteLn(o.Tag); o.Free; end."#
        ),
        &["12"]
    );
}

#[test]
fn virtual_override_tag_7() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public function Tag:Integer; virtual; end; TC=class(TB) function Tag:Integer; override; end; function TB.Tag:Integer; begin Result:=7; end; function TC.Tag:Integer; begin Result:=14; end; var o:TB; begin o:=TC.Create; WriteLn(o.Tag); o.Free; end."#
        ),
        &["14"]
    );
}

#[test]
fn virtual_override_tag_8() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public function Tag:Integer; virtual; end; TC=class(TB) function Tag:Integer; override; end; function TB.Tag:Integer; begin Result:=8; end; function TC.Tag:Integer; begin Result:=16; end; var o:TB; begin o:=TC.Create; WriteLn(o.Tag); o.Free; end."#
        ),
        &["16"]
    );
}

#[test]
fn virtual_override_tag_9() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public function Tag:Integer; virtual; end; TC=class(TB) function Tag:Integer; override; end; function TB.Tag:Integer; begin Result:=9; end; function TC.Tag:Integer; begin Result:=18; end; var o:TB; begin o:=TC.Create; WriteLn(o.Tag); o.Free; end."#
        ),
        &["18"]
    );
}

#[test]
fn virtual_override_tag_10() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public function Tag:Integer; virtual; end; TC=class(TB) function Tag:Integer; override; end; function TB.Tag:Integer; begin Result:=10; end; function TC.Tag:Integer; begin Result:=20; end; var o:TB; begin o:=TC.Create; WriteLn(o.Tag); o.Free; end."#
        ),
        &["20"]
    );
}

#[test]
fn virtual_override_tag_11() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public function Tag:Integer; virtual; end; TC=class(TB) function Tag:Integer; override; end; function TB.Tag:Integer; begin Result:=11; end; function TC.Tag:Integer; begin Result:=22; end; var o:TB; begin o:=TC.Create; WriteLn(o.Tag); o.Free; end."#
        ),
        &["22"]
    );
}

#[test]
fn virtual_override_tag_12() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public function Tag:Integer; virtual; end; TC=class(TB) function Tag:Integer; override; end; function TB.Tag:Integer; begin Result:=12; end; function TC.Tag:Integer; begin Result:=24; end; var o:TB; begin o:=TC.Create; WriteLn(o.Tag); o.Free; end."#
        ),
        &["24"]
    );
}

#[test]
fn virtual_override_tag_13() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public function Tag:Integer; virtual; end; TC=class(TB) function Tag:Integer; override; end; function TB.Tag:Integer; begin Result:=13; end; function TC.Tag:Integer; begin Result:=26; end; var o:TB; begin o:=TC.Create; WriteLn(o.Tag); o.Free; end."#
        ),
        &["26"]
    );
}

#[test]
fn virtual_override_tag_14() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public function Tag:Integer; virtual; end; TC=class(TB) function Tag:Integer; override; end; function TB.Tag:Integer; begin Result:=14; end; function TC.Tag:Integer; begin Result:=28; end; var o:TB; begin o:=TC.Create; WriteLn(o.Tag); o.Free; end."#
        ),
        &["28"]
    );
}
