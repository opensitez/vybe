/// Additional record operator overloads: add and equal.
use super::helpers::run_pascal;

#[test]
fn record_add_1() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Add(a,b:TR):TR; end; class operator TR.Add(a,b:TR):TR; begin Result.V:=a.V+b.V; end; var a,b,c:TR; begin a.V:=1; b.V:=2; c:=a+b; WriteLn(c.V); end."#
        ),
        &["3"]
    );
}

#[test]
fn record_add_2() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Add(a,b:TR):TR; end; class operator TR.Add(a,b:TR):TR; begin Result.V:=a.V+b.V; end; var a,b,c:TR; begin a.V:=2; b.V:=3; c:=a+b; WriteLn(c.V); end."#
        ),
        &["5"]
    );
}

#[test]
fn record_add_3() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Add(a,b:TR):TR; end; class operator TR.Add(a,b:TR):TR; begin Result.V:=a.V+b.V; end; var a,b,c:TR; begin a.V:=3; b.V:=4; c:=a+b; WriteLn(c.V); end."#
        ),
        &["7"]
    );
}

#[test]
fn record_add_4() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Add(a,b:TR):TR; end; class operator TR.Add(a,b:TR):TR; begin Result.V:=a.V+b.V; end; var a,b,c:TR; begin a.V:=4; b.V:=5; c:=a+b; WriteLn(c.V); end."#
        ),
        &["9"]
    );
}

#[test]
fn record_add_5() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Add(a,b:TR):TR; end; class operator TR.Add(a,b:TR):TR; begin Result.V:=a.V+b.V; end; var a,b,c:TR; begin a.V:=5; b.V:=6; c:=a+b; WriteLn(c.V); end."#
        ),
        &["11"]
    );
}

#[test]
fn record_add_6() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Add(a,b:TR):TR; end; class operator TR.Add(a,b:TR):TR; begin Result.V:=a.V+b.V; end; var a,b,c:TR; begin a.V:=6; b.V:=7; c:=a+b; WriteLn(c.V); end."#
        ),
        &["13"]
    );
}

#[test]
fn record_add_7() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Add(a,b:TR):TR; end; class operator TR.Add(a,b:TR):TR; begin Result.V:=a.V+b.V; end; var a,b,c:TR; begin a.V:=7; b.V:=8; c:=a+b; WriteLn(c.V); end."#
        ),
        &["15"]
    );
}

#[test]
fn record_add_8() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Add(a,b:TR):TR; end; class operator TR.Add(a,b:TR):TR; begin Result.V:=a.V+b.V; end; var a,b,c:TR; begin a.V:=8; b.V:=9; c:=a+b; WriteLn(c.V); end."#
        ),
        &["17"]
    );
}

#[test]
fn record_add_9() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Add(a,b:TR):TR; end; class operator TR.Add(a,b:TR):TR; begin Result.V:=a.V+b.V; end; var a,b,c:TR; begin a.V:=9; b.V:=10; c:=a+b; WriteLn(c.V); end."#
        ),
        &["19"]
    );
}

#[test]
fn record_add_10() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Add(a,b:TR):TR; end; class operator TR.Add(a,b:TR):TR; begin Result.V:=a.V+b.V; end; var a,b,c:TR; begin a.V:=10; b.V:=11; c:=a+b; WriteLn(c.V); end."#
        ),
        &["21"]
    );
}

#[test]
fn record_add_11() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Add(a,b:TR):TR; end; class operator TR.Add(a,b:TR):TR; begin Result.V:=a.V+b.V; end; var a,b,c:TR; begin a.V:=11; b.V:=12; c:=a+b; WriteLn(c.V); end."#
        ),
        &["23"]
    );
}

#[test]
fn record_add_12() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Add(a,b:TR):TR; end; class operator TR.Add(a,b:TR):TR; begin Result.V:=a.V+b.V; end; var a,b,c:TR; begin a.V:=12; b.V:=13; c:=a+b; WriteLn(c.V); end."#
        ),
        &["25"]
    );
}

#[test]
fn record_add_13() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Add(a,b:TR):TR; end; class operator TR.Add(a,b:TR):TR; begin Result.V:=a.V+b.V; end; var a,b,c:TR; begin a.V:=13; b.V:=14; c:=a+b; WriteLn(c.V); end."#
        ),
        &["27"]
    );
}

#[test]
fn record_add_14() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Add(a,b:TR):TR; end; class operator TR.Add(a,b:TR):TR; begin Result.V:=a.V+b.V; end; var a,b,c:TR; begin a.V:=14; b.V:=15; c:=a+b; WriteLn(c.V); end."#
        ),
        &["29"]
    );
}

#[test]
fn record_add_15() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Add(a,b:TR):TR; end; class operator TR.Add(a,b:TR):TR; begin Result.V:=a.V+b.V; end; var a,b,c:TR; begin a.V:=15; b.V:=16; c:=a+b; WriteLn(c.V); end."#
        ),
        &["31"]
    );
}

#[test]
fn record_add_16() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Add(a,b:TR):TR; end; class operator TR.Add(a,b:TR):TR; begin Result.V:=a.V+b.V; end; var a,b,c:TR; begin a.V:=16; b.V:=17; c:=a+b; WriteLn(c.V); end."#
        ),
        &["33"]
    );
}

#[test]
fn record_add_17() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Add(a,b:TR):TR; end; class operator TR.Add(a,b:TR):TR; begin Result.V:=a.V+b.V; end; var a,b,c:TR; begin a.V:=17; b.V:=18; c:=a+b; WriteLn(c.V); end."#
        ),
        &["35"]
    );
}

#[test]
fn record_add_18() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Add(a,b:TR):TR; end; class operator TR.Add(a,b:TR):TR; begin Result.V:=a.V+b.V; end; var a,b,c:TR; begin a.V:=18; b.V:=19; c:=a+b; WriteLn(c.V); end."#
        ),
        &["37"]
    );
}

#[test]
fn record_add_19() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Add(a,b:TR):TR; end; class operator TR.Add(a,b:TR):TR; begin Result.V:=a.V+b.V; end; var a,b,c:TR; begin a.V:=19; b.V:=20; c:=a+b; WriteLn(c.V); end."#
        ),
        &["39"]
    );
}

#[test]
fn record_add_20() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Add(a,b:TR):TR; end; class operator TR.Add(a,b:TR):TR; begin Result.V:=a.V+b.V; end; var a,b,c:TR; begin a.V:=20; b.V:=21; c:=a+b; WriteLn(c.V); end."#
        ),
        &["41"]
    );
}

#[test]
fn record_add_21() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Add(a,b:TR):TR; end; class operator TR.Add(a,b:TR):TR; begin Result.V:=a.V+b.V; end; var a,b,c:TR; begin a.V:=21; b.V:=22; c:=a+b; WriteLn(c.V); end."#
        ),
        &["43"]
    );
}

#[test]
fn record_equal_1() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Equal(a,b:TR):Boolean; end; class operator TR.Equal(a,b:TR):Boolean; begin Result:=a.V=b.V; end; var a,b:TR; begin a.V:=1; b.V:=1; if a=b then WriteLn(1) else WriteLn(0); end."#
        ),
        &["1"]
    );
}

#[test]
fn record_equal_2() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Equal(a,b:TR):Boolean; end; class operator TR.Equal(a,b:TR):Boolean; begin Result:=a.V=b.V; end; var a,b:TR; begin a.V:=2; b.V:=2; if a=b then WriteLn(1) else WriteLn(0); end."#
        ),
        &["1"]
    );
}

#[test]
fn record_equal_3() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Equal(a,b:TR):Boolean; end; class operator TR.Equal(a,b:TR):Boolean; begin Result:=a.V=b.V; end; var a,b:TR; begin a.V:=3; b.V:=3; if a=b then WriteLn(1) else WriteLn(0); end."#
        ),
        &["1"]
    );
}

#[test]
fn record_equal_4() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Equal(a,b:TR):Boolean; end; class operator TR.Equal(a,b:TR):Boolean; begin Result:=a.V=b.V; end; var a,b:TR; begin a.V:=4; b.V:=4; if a=b then WriteLn(1) else WriteLn(0); end."#
        ),
        &["1"]
    );
}

#[test]
fn record_equal_5() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Equal(a,b:TR):Boolean; end; class operator TR.Equal(a,b:TR):Boolean; begin Result:=a.V=b.V; end; var a,b:TR; begin a.V:=5; b.V:=5; if a=b then WriteLn(1) else WriteLn(0); end."#
        ),
        &["1"]
    );
}

#[test]
fn record_equal_6() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Equal(a,b:TR):Boolean; end; class operator TR.Equal(a,b:TR):Boolean; begin Result:=a.V=b.V; end; var a,b:TR; begin a.V:=6; b.V:=6; if a=b then WriteLn(1) else WriteLn(0); end."#
        ),
        &["1"]
    );
}

#[test]
fn record_equal_7() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Equal(a,b:TR):Boolean; end; class operator TR.Equal(a,b:TR):Boolean; begin Result:=a.V=b.V; end; var a,b:TR; begin a.V:=7; b.V:=7; if a=b then WriteLn(1) else WriteLn(0); end."#
        ),
        &["1"]
    );
}

#[test]
fn record_equal_8() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Equal(a,b:TR):Boolean; end; class operator TR.Equal(a,b:TR):Boolean; begin Result:=a.V=b.V; end; var a,b:TR; begin a.V:=8; b.V:=8; if a=b then WriteLn(1) else WriteLn(0); end."#
        ),
        &["1"]
    );
}

#[test]
fn record_equal_9() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Equal(a,b:TR):Boolean; end; class operator TR.Equal(a,b:TR):Boolean; begin Result:=a.V=b.V; end; var a,b:TR; begin a.V:=9; b.V:=9; if a=b then WriteLn(1) else WriteLn(0); end."#
        ),
        &["1"]
    );
}

#[test]
fn record_equal_10() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Equal(a,b:TR):Boolean; end; class operator TR.Equal(a,b:TR):Boolean; begin Result:=a.V=b.V; end; var a,b:TR; begin a.V:=10; b.V:=10; if a=b then WriteLn(1) else WriteLn(0); end."#
        ),
        &["1"]
    );
}

#[test]
fn record_equal_11() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Equal(a,b:TR):Boolean; end; class operator TR.Equal(a,b:TR):Boolean; begin Result:=a.V=b.V; end; var a,b:TR; begin a.V:=11; b.V:=11; if a=b then WriteLn(1) else WriteLn(0); end."#
        ),
        &["1"]
    );
}

#[test]
fn record_equal_12() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Equal(a,b:TR):Boolean; end; class operator TR.Equal(a,b:TR):Boolean; begin Result:=a.V=b.V; end; var a,b:TR; begin a.V:=12; b.V:=12; if a=b then WriteLn(1) else WriteLn(0); end."#
        ),
        &["1"]
    );
}

#[test]
fn record_equal_13() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Equal(a,b:TR):Boolean; end; class operator TR.Equal(a,b:TR):Boolean; begin Result:=a.V=b.V; end; var a,b:TR; begin a.V:=13; b.V:=13; if a=b then WriteLn(1) else WriteLn(0); end."#
        ),
        &["1"]
    );
}

#[test]
fn record_equal_14() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Equal(a,b:TR):Boolean; end; class operator TR.Equal(a,b:TR):Boolean; begin Result:=a.V=b.V; end; var a,b:TR; begin a.V:=14; b.V:=14; if a=b then WriteLn(1) else WriteLn(0); end."#
        ),
        &["1"]
    );
}

#[test]
fn record_equal_15() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Equal(a,b:TR):Boolean; end; class operator TR.Equal(a,b:TR):Boolean; begin Result:=a.V=b.V; end; var a,b:TR; begin a.V:=15; b.V:=15; if a=b then WriteLn(1) else WriteLn(0); end."#
        ),
        &["1"]
    );
}

#[test]
fn record_equal_16() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Equal(a,b:TR):Boolean; end; class operator TR.Equal(a,b:TR):Boolean; begin Result:=a.V=b.V; end; var a,b:TR; begin a.V:=16; b.V:=16; if a=b then WriteLn(1) else WriteLn(0); end."#
        ),
        &["1"]
    );
}

#[test]
fn record_equal_17() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Equal(a,b:TR):Boolean; end; class operator TR.Equal(a,b:TR):Boolean; begin Result:=a.V=b.V; end; var a,b:TR; begin a.V:=17; b.V:=17; if a=b then WriteLn(1) else WriteLn(0); end."#
        ),
        &["1"]
    );
}

#[test]
fn record_equal_18() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Equal(a,b:TR):Boolean; end; class operator TR.Equal(a,b:TR):Boolean; begin Result:=a.V=b.V; end; var a,b:TR; begin a.V:=18; b.V:=18; if a=b then WriteLn(1) else WriteLn(0); end."#
        ),
        &["1"]
    );
}

#[test]
fn record_equal_19() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Equal(a,b:TR):Boolean; end; class operator TR.Equal(a,b:TR):Boolean; begin Result:=a.V=b.V; end; var a,b:TR; begin a.V:=19; b.V:=19; if a=b then WriteLn(1) else WriteLn(0); end."#
        ),
        &["1"]
    );
}

#[test]
fn record_equal_20() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Equal(a,b:TR):Boolean; end; class operator TR.Equal(a,b:TR):Boolean; begin Result:=a.V=b.V; end; var a,b:TR; begin a.V:=20; b.V:=20; if a=b then WriteLn(1) else WriteLn(0); end."#
        ),
        &["1"]
    );
}

#[test]
fn record_equal_21() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; class operator Equal(a,b:TR):Boolean; end; class operator TR.Equal(a,b:TR):Boolean; begin Result:=a.V=b.V; end; var a,b:TR; begin a.V:=21; b.V:=21; if a=b then WriteLn(1) else WriteLn(0); end."#
        ),
        &["1"]
    );
}
