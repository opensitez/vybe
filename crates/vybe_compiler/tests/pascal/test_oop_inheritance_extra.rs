/// Deeper inheritance trees: multi-level chains, overrides, constructors.
use super::helpers::run_pascal;

#[test]
fn inheritance_chain_4_levels() {
    assert_eq!(
        run_pascal(
            r#"program T; type TL1=class public function L:Integer; virtual; end; TL2=class(TL1) function L:Integer; override; end; TL3=class(TL2) function L:Integer; override; end;  function TL1.L:Integer; begin Result:=1; end; function TL2.L:Integer; begin Result:=2; end; function TL3.L:Integer; begin Result:=3; end;  var o:TL1; begin o:=TL3.Create; WriteLn(o.L); o.Free; end."#
        ),
        &["3"]
    );
}

#[test]
fn inheritance_chain_5_levels() {
    assert_eq!(
        run_pascal(
            r#"program T; type TL1=class public function L:Integer; virtual; end; TL2=class(TL1) function L:Integer; override; end; TL3=class(TL2) function L:Integer; override; end; TL4=class(TL3) function L:Integer; override; end;  function TL1.L:Integer; begin Result:=1; end; function TL2.L:Integer; begin Result:=2; end; function TL3.L:Integer; begin Result:=3; end; function TL4.L:Integer; begin Result:=4; end;  var o:TL1; begin o:=TL4.Create; WriteLn(o.L); o.Free; end."#
        ),
        &["4"]
    );
}

#[test]
fn inheritance_chain_6_levels() {
    assert_eq!(
        run_pascal(
            r#"program T; type TL1=class public function L:Integer; virtual; end; TL2=class(TL1) function L:Integer; override; end; TL3=class(TL2) function L:Integer; override; end; TL4=class(TL3) function L:Integer; override; end; TL5=class(TL4) function L:Integer; override; end;  function TL1.L:Integer; begin Result:=1; end; function TL2.L:Integer; begin Result:=2; end; function TL3.L:Integer; begin Result:=3; end; function TL4.L:Integer; begin Result:=4; end; function TL5.L:Integer; begin Result:=5; end;  var o:TL1; begin o:=TL5.Create; WriteLn(o.L); o.Free; end."#
        ),
        &["5"]
    );
}

#[test]
fn inheritance_chain_7_levels() {
    assert_eq!(
        run_pascal(
            r#"program T; type TL1=class public function L:Integer; virtual; end; TL2=class(TL1) function L:Integer; override; end; TL3=class(TL2) function L:Integer; override; end; TL4=class(TL3) function L:Integer; override; end; TL5=class(TL4) function L:Integer; override; end; TL6=class(TL5) function L:Integer; override; end;  function TL1.L:Integer; begin Result:=1; end; function TL2.L:Integer; begin Result:=2; end; function TL3.L:Integer; begin Result:=3; end; function TL4.L:Integer; begin Result:=4; end; function TL5.L:Integer; begin Result:=5; end; function TL6.L:Integer; begin Result:=6; end;  var o:TL1; begin o:=TL6.Create; WriteLn(o.L); o.Free; end."#
        ),
        &["6"]
    );
}

#[test]
fn inheritance_chain_8_levels() {
    assert_eq!(
        run_pascal(
            r#"program T; type TL1=class public function L:Integer; virtual; end; TL2=class(TL1) function L:Integer; override; end; TL3=class(TL2) function L:Integer; override; end; TL4=class(TL3) function L:Integer; override; end; TL5=class(TL4) function L:Integer; override; end; TL6=class(TL5) function L:Integer; override; end; TL7=class(TL6) function L:Integer; override; end;  function TL1.L:Integer; begin Result:=1; end; function TL2.L:Integer; begin Result:=2; end; function TL3.L:Integer; begin Result:=3; end; function TL4.L:Integer; begin Result:=4; end; function TL5.L:Integer; begin Result:=5; end; function TL6.L:Integer; begin Result:=6; end; function TL7.L:Integer; begin Result:=7; end;  var o:TL1; begin o:=TL7.Create; WriteLn(o.L); o.Free; end."#
        ),
        &["7"]
    );
}

#[test]
fn inherited_field_depth_1() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBase=class public V:Integer; end; T1=class(TBase);  var o:T1; begin o:=T1.Create; o.V:=3; WriteLn(o.V); o.Free; end."#
        ),
        &["3"]
    );
}

#[test]
fn inherited_field_depth_2() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBase=class public V:Integer; end; T1=class(TBase); T2=class(T1);  var o:T2; begin o:=T2.Create; o.V:=6; WriteLn(o.V); o.Free; end."#
        ),
        &["6"]
    );
}

#[test]
fn inherited_field_depth_3() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBase=class public V:Integer; end; T1=class(TBase); T2=class(T1); T3=class(T2);  var o:T3; begin o:=T3.Create; o.V:=9; WriteLn(o.V); o.Free; end."#
        ),
        &["9"]
    );
}

#[test]
fn inherited_field_depth_4() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBase=class public V:Integer; end; T1=class(TBase); T2=class(T1); T3=class(T2); T4=class(T3);  var o:T4; begin o:=T4.Create; o.V:=12; WriteLn(o.V); o.Free; end."#
        ),
        &["12"]
    );
}

#[test]
fn inherited_field_depth_5() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBase=class public V:Integer; end; T1=class(TBase); T2=class(T1); T3=class(T2); T4=class(T3); T5=class(T4);  var o:T5; begin o:=T5.Create; o.V:=15; WriteLn(o.V); o.Free; end."#
        ),
        &["15"]
    );
}

#[test]
fn inherited_field_depth_6() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBase=class public V:Integer; end; T1=class(TBase); T2=class(T1); T3=class(T2); T4=class(T3); T5=class(T4); T6=class(T5);  var o:T6; begin o:=T6.Create; o.V:=18; WriteLn(o.V); o.Free; end."#
        ),
        &["18"]
    );
}

#[test]
fn inherited_field_depth_7() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBase=class public V:Integer; end; T1=class(TBase); T2=class(T1); T3=class(T2); T4=class(T3); T5=class(T4); T6=class(T5); T7=class(T6);  var o:T7; begin o:=T7.Create; o.V:=21; WriteLn(o.V); o.Free; end."#
        ),
        &["21"]
    );
}

#[test]
fn inherited_field_depth_8() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBase=class public V:Integer; end; T1=class(TBase); T2=class(T1); T3=class(T2); T4=class(T3); T5=class(T4); T6=class(T5); T7=class(T6); T8=class(T7);  var o:T8; begin o:=T8.Create; o.V:=24; WriteLn(o.V); o.Free; end."#
        ),
        &["24"]
    );
}

#[test]
fn inherited_field_depth_9() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBase=class public V:Integer; end; T1=class(TBase); T2=class(T1); T3=class(T2); T4=class(T3); T5=class(T4); T6=class(T5); T7=class(T6); T8=class(T7); T9=class(T8);  var o:T9; begin o:=T9.Create; o.V:=27; WriteLn(o.V); o.Free; end."#
        ),
        &["27"]
    );
}

#[test]
fn inherited_field_depth_10() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBase=class public V:Integer; end; T1=class(TBase); T2=class(T1); T3=class(T2); T4=class(T3); T5=class(T4); T6=class(T5); T7=class(T6); T8=class(T7); T9=class(T8); T10=class(T9);  var o:T10; begin o:=T10.Create; o.V:=30; WriteLn(o.V); o.Free; end."#
        ),
        &["30"]
    );
}

#[test]
fn override_inherited_plus_1() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public function F:Integer; virtual; end; TC=class(TB) function F:Integer; override; end; function TB.F:Integer; begin Result:=1; end; function TC.F:Integer; begin Result:=inherited F+1; end; var c:TC; begin c:=TC.Create; WriteLn(c.F); c.Free; end."#
        ),
        &["2"]
    );
}

#[test]
fn override_inherited_plus_2() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public function F:Integer; virtual; end; TC=class(TB) function F:Integer; override; end; function TB.F:Integer; begin Result:=2; end; function TC.F:Integer; begin Result:=inherited F+2; end; var c:TC; begin c:=TC.Create; WriteLn(c.F); c.Free; end."#
        ),
        &["4"]
    );
}

#[test]
fn override_inherited_plus_3() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public function F:Integer; virtual; end; TC=class(TB) function F:Integer; override; end; function TB.F:Integer; begin Result:=3; end; function TC.F:Integer; begin Result:=inherited F+3; end; var c:TC; begin c:=TC.Create; WriteLn(c.F); c.Free; end."#
        ),
        &["6"]
    );
}

#[test]
fn override_inherited_plus_4() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public function F:Integer; virtual; end; TC=class(TB) function F:Integer; override; end; function TB.F:Integer; begin Result:=4; end; function TC.F:Integer; begin Result:=inherited F+4; end; var c:TC; begin c:=TC.Create; WriteLn(c.F); c.Free; end."#
        ),
        &["8"]
    );
}

#[test]
fn override_inherited_plus_5() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public function F:Integer; virtual; end; TC=class(TB) function F:Integer; override; end; function TB.F:Integer; begin Result:=5; end; function TC.F:Integer; begin Result:=inherited F+5; end; var c:TC; begin c:=TC.Create; WriteLn(c.F); c.Free; end."#
        ),
        &["10"]
    );
}

#[test]
fn override_inherited_plus_6() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public function F:Integer; virtual; end; TC=class(TB) function F:Integer; override; end; function TB.F:Integer; begin Result:=6; end; function TC.F:Integer; begin Result:=inherited F+6; end; var c:TC; begin c:=TC.Create; WriteLn(c.F); c.Free; end."#
        ),
        &["12"]
    );
}

#[test]
fn override_inherited_plus_7() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public function F:Integer; virtual; end; TC=class(TB) function F:Integer; override; end; function TB.F:Integer; begin Result:=7; end; function TC.F:Integer; begin Result:=inherited F+7; end; var c:TC; begin c:=TC.Create; WriteLn(c.F); c.Free; end."#
        ),
        &["14"]
    );
}

#[test]
fn override_inherited_plus_8() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public function F:Integer; virtual; end; TC=class(TB) function F:Integer; override; end; function TB.F:Integer; begin Result:=8; end; function TC.F:Integer; begin Result:=inherited F+8; end; var c:TC; begin c:=TC.Create; WriteLn(c.F); c.Free; end."#
        ),
        &["16"]
    );
}

#[test]
fn override_inherited_plus_9() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public function F:Integer; virtual; end; TC=class(TB) function F:Integer; override; end; function TB.F:Integer; begin Result:=9; end; function TC.F:Integer; begin Result:=inherited F+9; end; var c:TC; begin c:=TC.Create; WriteLn(c.F); c.Free; end."#
        ),
        &["18"]
    );
}

#[test]
fn override_inherited_plus_10() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public function F:Integer; virtual; end; TC=class(TB) function F:Integer; override; end; function TB.F:Integer; begin Result:=10; end; function TC.F:Integer; begin Result:=inherited F+10; end; var c:TC; begin c:=TC.Create; WriteLn(c.F); c.Free; end."#
        ),
        &["20"]
    );
}

#[test]
fn sibling_override_1_2() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public function N:Integer; virtual; end; TA=class(TB) function N:Integer; override; end; TC=class(TB) function N:Integer; override; end; function TB.N:Integer; begin Result:=0; end; function TA.N:Integer; begin Result:=1; end; function TC.N:Integer; begin Result:=2; end; var x:TA; y:TC; begin x:=TA.Create; y:=TC.Create; WriteLn(x.N); WriteLn(y.N); x.Free; y.Free; end."#
        ),
        &["1", "2"]
    );
}

#[test]
fn sibling_override_3_4() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public function N:Integer; virtual; end; TA=class(TB) function N:Integer; override; end; TC=class(TB) function N:Integer; override; end; function TB.N:Integer; begin Result:=0; end; function TA.N:Integer; begin Result:=3; end; function TC.N:Integer; begin Result:=4; end; var x:TA; y:TC; begin x:=TA.Create; y:=TC.Create; WriteLn(x.N); WriteLn(y.N); x.Free; y.Free; end."#
        ),
        &["3", "4"]
    );
}

#[test]
fn sibling_override_5_6() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public function N:Integer; virtual; end; TA=class(TB) function N:Integer; override; end; TC=class(TB) function N:Integer; override; end; function TB.N:Integer; begin Result:=0; end; function TA.N:Integer; begin Result:=5; end; function TC.N:Integer; begin Result:=6; end; var x:TA; y:TC; begin x:=TA.Create; y:=TC.Create; WriteLn(x.N); WriteLn(y.N); x.Free; y.Free; end."#
        ),
        &["5", "6"]
    );
}

#[test]
fn sibling_override_7_8() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public function N:Integer; virtual; end; TA=class(TB) function N:Integer; override; end; TC=class(TB) function N:Integer; override; end; function TB.N:Integer; begin Result:=0; end; function TA.N:Integer; begin Result:=7; end; function TC.N:Integer; begin Result:=8; end; var x:TA; y:TC; begin x:=TA.Create; y:=TC.Create; WriteLn(x.N); WriteLn(y.N); x.Free; y.Free; end."#
        ),
        &["7", "8"]
    );
}

#[test]
fn sibling_override_9_10() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public function N:Integer; virtual; end; TA=class(TB) function N:Integer; override; end; TC=class(TB) function N:Integer; override; end; function TB.N:Integer; begin Result:=0; end; function TA.N:Integer; begin Result:=9; end; function TC.N:Integer; begin Result:=10; end; var x:TA; y:TC; begin x:=TA.Create; y:=TC.Create; WriteLn(x.N); WriteLn(y.N); x.Free; y.Free; end."#
        ),
        &["9", "10"]
    );
}

#[test]
fn sibling_override_2_5() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public function N:Integer; virtual; end; TA=class(TB) function N:Integer; override; end; TC=class(TB) function N:Integer; override; end; function TB.N:Integer; begin Result:=0; end; function TA.N:Integer; begin Result:=2; end; function TC.N:Integer; begin Result:=5; end; var x:TA; y:TC; begin x:=TA.Create; y:=TC.Create; WriteLn(x.N); WriteLn(y.N); x.Free; y.Free; end."#
        ),
        &["2", "5"]
    );
}

#[test]
fn sibling_override_4_9() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public function N:Integer; virtual; end; TA=class(TB) function N:Integer; override; end; TC=class(TB) function N:Integer; override; end; function TB.N:Integer; begin Result:=0; end; function TA.N:Integer; begin Result:=4; end; function TC.N:Integer; begin Result:=9; end; var x:TA; y:TC; begin x:=TA.Create; y:=TC.Create; WriteLn(x.N); WriteLn(y.N); x.Free; y.Free; end."#
        ),
        &["4", "9"]
    );
}

#[test]
fn sibling_override_6_1() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public function N:Integer; virtual; end; TA=class(TB) function N:Integer; override; end; TC=class(TB) function N:Integer; override; end; function TB.N:Integer; begin Result:=0; end; function TA.N:Integer; begin Result:=6; end; function TC.N:Integer; begin Result:=1; end; var x:TA; y:TC; begin x:=TA.Create; y:=TC.Create; WriteLn(x.N); WriteLn(y.N); x.Free; y.Free; end."#
        ),
        &["6", "1"]
    );
}

#[test]
fn sibling_override_8_3() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public function N:Integer; virtual; end; TA=class(TB) function N:Integer; override; end; TC=class(TB) function N:Integer; override; end; function TB.N:Integer; begin Result:=0; end; function TA.N:Integer; begin Result:=8; end; function TC.N:Integer; begin Result:=3; end; var x:TA; y:TC; begin x:=TA.Create; y:=TC.Create; WriteLn(x.N); WriteLn(y.N); x.Free; y.Free; end."#
        ),
        &["8", "3"]
    );
}

#[test]
fn sibling_override_10_7() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=class public function N:Integer; virtual; end; TA=class(TB) function N:Integer; override; end; TC=class(TB) function N:Integer; override; end; function TB.N:Integer; begin Result:=0; end; function TA.N:Integer; begin Result:=10; end; function TC.N:Integer; begin Result:=7; end; var x:TA; y:TC; begin x:=TA.Create; y:=TC.Create; WriteLn(x.N); WriteLn(y.N); x.Free; y.Free; end."#
        ),
        &["10", "7"]
    );
}

#[test]
fn constructor_chain_add_1() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBase=class public Value:Integer; constructor Create(v:Integer); end; TChild=class(TBase) constructor Create(v:Integer); end; constructor TBase.Create(v:Integer); begin Value:=v; end; constructor TChild.Create(v:Integer); begin inherited Create(v+1); end; var c:TChild; begin c:=TChild.Create(10); WriteLn(c.Value); c.Free; end."#
        ),
        &["11"]
    );
}

#[test]
fn constructor_chain_add_2() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBase=class public Value:Integer; constructor Create(v:Integer); end; TChild=class(TBase) constructor Create(v:Integer); end; constructor TBase.Create(v:Integer); begin Value:=v; end; constructor TChild.Create(v:Integer); begin inherited Create(v+2); end; var c:TChild; begin c:=TChild.Create(10); WriteLn(c.Value); c.Free; end."#
        ),
        &["12"]
    );
}

#[test]
fn constructor_chain_add_3() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBase=class public Value:Integer; constructor Create(v:Integer); end; TChild=class(TBase) constructor Create(v:Integer); end; constructor TBase.Create(v:Integer); begin Value:=v; end; constructor TChild.Create(v:Integer); begin inherited Create(v+3); end; var c:TChild; begin c:=TChild.Create(10); WriteLn(c.Value); c.Free; end."#
        ),
        &["13"]
    );
}

#[test]
fn constructor_chain_add_4() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBase=class public Value:Integer; constructor Create(v:Integer); end; TChild=class(TBase) constructor Create(v:Integer); end; constructor TBase.Create(v:Integer); begin Value:=v; end; constructor TChild.Create(v:Integer); begin inherited Create(v+4); end; var c:TChild; begin c:=TChild.Create(10); WriteLn(c.Value); c.Free; end."#
        ),
        &["14"]
    );
}

#[test]
fn constructor_chain_add_5() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBase=class public Value:Integer; constructor Create(v:Integer); end; TChild=class(TBase) constructor Create(v:Integer); end; constructor TBase.Create(v:Integer); begin Value:=v; end; constructor TChild.Create(v:Integer); begin inherited Create(v+5); end; var c:TChild; begin c:=TChild.Create(10); WriteLn(c.Value); c.Free; end."#
        ),
        &["15"]
    );
}

#[test]
fn constructor_chain_add_6() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBase=class public Value:Integer; constructor Create(v:Integer); end; TChild=class(TBase) constructor Create(v:Integer); end; constructor TBase.Create(v:Integer); begin Value:=v; end; constructor TChild.Create(v:Integer); begin inherited Create(v+6); end; var c:TChild; begin c:=TChild.Create(10); WriteLn(c.Value); c.Free; end."#
        ),
        &["16"]
    );
}

#[test]
fn constructor_chain_add_7() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBase=class public Value:Integer; constructor Create(v:Integer); end; TChild=class(TBase) constructor Create(v:Integer); end; constructor TBase.Create(v:Integer); begin Value:=v; end; constructor TChild.Create(v:Integer); begin inherited Create(v+7); end; var c:TChild; begin c:=TChild.Create(10); WriteLn(c.Value); c.Free; end."#
        ),
        &["17"]
    );
}

