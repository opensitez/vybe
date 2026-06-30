/// Class inheritance patterns: extends, overrides, protected, virtual constructors.
use super::helpers::run_pascal;

#[test]
fn inherited_field_visible_in_child() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBase=class public Value:Integer; end; TChild=class(TBase); var c:TChild; begin c:=TChild.Create; c.Value:=8; WriteLn(c.Value); c.Free; end."#
        ),
        &["8"]
    );
}

#[test]
fn inherited_method_called_on_child() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBase=class public function Tag:string; virtual; end; TChild=class(TBase) function Tag:string; override; end; function TBase.Tag:string; begin Result:='base'; end; function TChild.Tag:string; begin Result:='child'; end; var c:TChild; begin c:=TChild.Create; WriteLn(c.Tag); c.Free; end."#
        ),
        &["child"]
    );
}

#[test]
fn base_reference_virtual_dispatch() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBase=class public function Id:Integer; virtual; end; TChild=class(TBase) function Id:Integer; override; end; function TBase.Id:Integer; begin Result:=1; end; function TChild.Id:Integer; begin Result:=2; end; var b:TBase; begin b:=TChild.Create; WriteLn(b.Id); b.Free; end."#
        ),
        &["2"]
    );
}

#[test]
fn inherited_constructor_chain() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBase=class public Value:Integer; constructor Create(v:Integer); end; TChild=class(TBase) constructor Create(v:Integer); end; constructor TBase.Create(v:Integer); begin Value:=v; end; constructor TChild.Create(v:Integer); begin inherited Create(v+1); end; var c:TChild; begin c:=TChild.Create(4); WriteLn(c.Value); c.Free; end."#
        ),
        &["5"]
    );
}

#[test]
fn protected_member_visible_in_descendant() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBase=class strict protected F:Integer; public function Get:Integer; end; TChild=class(TBase) procedure Set(v:Integer); end; function TBase.Get:Integer; begin Result:=F; end; procedure TChild.Set(v:Integer); begin F:=v; end; var c:TChild; begin c:=TChild.Create; c.Set(6); WriteLn(c.Get); c.Free; end."#
        ),
        &["6"]
    );
}

#[test]
fn override_calls_inherited() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBase=class public function Twice(n:Integer):Integer; virtual; end; TChild=class(TBase) function Twice(n:Integer):Integer; override; end; function TBase.Twice(n:Integer):Integer; begin Result:=n*2; end; function TChild.Twice(n:Integer):Integer; begin Result:=inherited Twice(n)+1; end; var c:TChild; begin c:=TChild.Create; WriteLn(c.Twice(3)); c.Free; end."#
        ),
        &["7"]
    );
}

#[test]
fn deep_inheritance_three_levels() {
    assert_eq!(
        run_pascal(
            r#"program T; type TA=class public function L:Integer; virtual; end; TB=class(TA) function L:Integer; override; end; TC=class(TB) function L:Integer; override; end; function TA.L:Integer; begin Result:=1; end; function TB.L:Integer; begin Result:=2; end; function TC.L:Integer; begin Result:=3; end; var o:TA; begin o:=TC.Create; WriteLn(o.L); o.Free; end."#
        ),
        &["3"]
    );
}

#[test]
fn sibling_classes_do_not_share_overrides() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBase=class public function N:Integer; virtual; end; TA=class(TBase) function N:Integer; override; end; TB=class(TBase) function N:Integer; override; end; function TBase.N:Integer; begin Result:=0; end; function TA.N:Integer; begin Result:=1; end; function TB.N:Integer; begin Result:=2; end; var a:TA; b:TB; begin a:=TA.Create; b:=TB.Create; WriteLn(a.N); WriteLn(b.N); a.Free; b.Free; end."#
        ),
        &["1", "2"]
    );
}

#[test]
fn class_type_check_via_is_operator() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBase=class end; TChild=class(TBase); var b:TBase; begin b:=TChild.Create; if b is TChild then WriteLn('yes'); b.Free; end."#
        ),
        &["yes"]
    );
}

#[test]
fn class_cast_as_child() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBase=class public V:Integer; end; TChild=class(TBase); var b:TBase; c:TChild; begin b:=TChild.Create; b.V:=9; c:=TChild(b); WriteLn(c.V); c.Free; end."#
        ),
        &["9"]
    );
}

#[test]
fn inherited_destructor_virtual_chain() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBase=class public class var DtorCount:Integer; destructor Destroy; virtual; end; TChild=class(TBase) destructor Destroy; override; end; class var TBase.DtorCount:Integer; destructor TBase.Destroy; begin Inc(DtorCount); inherited; end; destructor TChild.Destroy; begin inherited; end; var c:TChild; begin TBase.DtorCount:=0; c:=TChild.Create; c.Free; WriteLn(TBase.DtorCount); end."#
        ),
        &["1"]
    );
}

#[test]
fn abstract_parent_concrete_child() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBase=class abstract public function Run:Integer; virtual; abstract; end; TChild=class(TBase) function Run:Integer; override; end; function TChild.Run:Integer; begin Result:=5; end; var b:TBase; begin b:=TChild.Create; WriteLn(b.Run); b.Free; end."#
        ),
        &["5"]
    );
}

#[test]
fn reintroduce_method_hides_parent() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBase=class public function Name:string; virtual; end; TChild=class(TBase) function Name:string; reintroduce; end; function TBase.Name:string; begin Result:='base'; end; function TChild.Name:string; begin Result:='child'; end; var c:TChild; begin c:=TChild.Create; WriteLn(c.Name); c.Free; end."#
        ),
        &["child"]
    );
}

#[test]
fn inherited_class_var_shared() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBase=class public class var Hits:Integer; end; TChild=class(TBase); class var TBase.Hits:Integer; begin TBase.Hits:=0; Inc(TChild.Hits); WriteLn(TBase.Hits); end."#
        ),
        &["1"]
    );
}

#[test]
fn override_with_different_visibility_not_allowed_use_public() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBase=class public procedure Ping; virtual; end; TChild=class(TBase) public procedure Ping; override; end; procedure TBase.Ping; begin WriteLn('base'); end; procedure TChild.Ping; begin WriteLn('child'); end; var b:TBase; begin b:=TChild.Create; b.Ping; b.Free; end."#
        ),
        &["child"]
    );
}

#[test]
fn parent_method_not_overridden_keeps_base_behavior() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBase=class public function Val:Integer; virtual; end; TChild=class(TBase) function Other:Integer; end; function TBase.Val:Integer; begin Result:=10; end; function TChild.Other:Integer; begin Result:=20; end; var c:TChild; begin c:=TChild.Create; WriteLn(c.Val); c.Free; end."#
        ),
        &["10"]
    );
}

#[test]
fn interface_inheritance_on_classes() {
    assert_eq!(
        run_pascal(
            r#"program T; type ICore=interface function Core:Integer; end; TBase=class(TInterfacedObject,ICore) function Core:Integer; virtual; end; TChild=class(TBase) function Core:Integer; override; end; function TBase.Core:Integer; begin Result:=1; end; function TChild.Core:Integer; begin Result:=2; end; var i:ICore; begin i:=TChild.Create; WriteLn(i.Core); end."#
        ),
        &["2"]
    );
}

#[test]
fn field_shadowing_in_child() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBase=class public X:Integer; end; TChild=class(TBase) public Y:Integer; end; var c:TChild; begin c:=TChild.Create; c.X:=1; c.Y:=2; WriteLn(c.X+c.Y); c.Free; end."#
        ),
        &["3"]
    );
}

#[test]
fn grandchild_overrides_parent_virtual() {
    assert_eq!(
        run_pascal(
            r#"program T; type TA=class public function F:Integer; virtual; end; TB=class(TA) end; TC=class(TB) function F:Integer; override; end; function TA.F:Integer; begin Result:=1; end; function TC.F:Integer; begin Result:=3; end; var o:TA; begin o:=TC.Create; WriteLn(o.F); o.Free; end."#
        ),
        &["3"]
    );
}

#[test]
fn inherited_property_reader() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBase=class private F:Integer; public property N:Integer read F write F; end; TChild=class(TBase); var c:TChild; begin c:=TChild.Create; c.N:=12; WriteLn(c.N); c.Free; end."#
        ),
        &["12"]
    );
}
