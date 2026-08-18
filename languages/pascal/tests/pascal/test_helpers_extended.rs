/// Record helpers, class helpers, and extended helper patterns.
use super::helpers::run_pascal;

#[test]
fn record_helper_adds_method_to_record() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPoint=record X,Y:Integer; end; TPointHelper=record helper for TPoint function Sum:Integer; end; function TPointHelper.Sum:Integer; begin Result:=X+Y; end; var p:TPoint; begin p.X:=2; p.Y:=3; WriteLn(p.Sum); end."#
        ),
        &["5"]
    );
}

#[test]
fn class_helper_extends_class_api() {
    assert_eq!(
        run_pascal(
            r#"program T; type TList=class public Count:Integer; end; TListHelper=class helper for TList function IsEmpty:Boolean; end; function TListHelper.IsEmpty:Boolean; begin Result:=Count=0; end; var L:TList; begin L:=TList.Create; WriteLn(L.IsEmpty); L.Count:=1; WriteLn(L.IsEmpty); L.Free; end."#
        ),
        &["true", "false"]
    );
}

#[test]
fn integer_helper_is_positive() {
    assert_eq!(
        run_pascal(
            r#"program T; type TIntHelper=record helper for Integer function IsPositive:Boolean; end; function TIntHelper.IsPositive:Boolean; begin Result:=Self>0; end; var n:Integer; begin n:=5; WriteLn(n.IsPositive); n:=-1; WriteLn(n.IsPositive); end."#
        ),
        &["true", "false"]
    );
}

#[test]
fn string_helper_starts_with_prefix() {
    assert_eq!(
        run_pascal(
            r#"program T; type TStrHelper=record helper for String function Starts(prefix:string):Boolean; end; function TStrHelper.Starts(prefix:string):Boolean; begin Result:=Copy(Self,1,Length(prefix))=prefix; end; var s:string; begin s:='hello'; WriteLn(s.Starts('he')); WriteLn(s.Starts('zz')); end."#
        ),
        &["true", "false"]
    );
}

#[test]
fn char_helper_is_digit() {
    assert_eq!(
        run_pascal(
            r#"program T; type TCharHelper=record helper for Char function IsDigit:Boolean; end; function TCharHelper.IsDigit:Boolean; begin Result:=(Self>='0') and (Self<='9'); end; begin WriteLn('7'.IsDigit); WriteLn('a'.IsDigit); end."#
        ),
        &["true", "false"]
    );
}

#[test]
fn double_helper_abs_value() {
    assert_eq!(
        run_pascal(
            r#"program T; type TDoubleHelper=record helper for Double function AbsVal:Double; end; function TDoubleHelper.AbsVal:Double; begin if Self<0 then Result:=-Self else Result:=Self; end; var d:Double; begin d:=-2.5; WriteLn(d.AbsVal); end."#
        ),
        &["2.5"]
    );
}

#[test]
fn boolean_helper_toggle() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBoolHelper=record helper for Boolean function Toggle:Boolean; end; function TBoolHelper.Toggle:Boolean; begin Result:=not Self; end; var b:Boolean; begin b:=true; b:=b.Toggle; WriteLn(b); end."#
        ),
        &["FALSE"]
    );
}

#[test]
fn record_helper_chain_two_calls() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; end; TRHelper=record helper for TR function IncV:TR; function DoubleV:TR; end; function TRHelper.IncV:TR; begin Result:=Self; Result.V:=V+1; end; function TRHelper.DoubleV:TR; begin Result:=Self; Result.V:=V*2; end; var r:TR; begin r.V:=3; r:=r.IncV.DoubleV; WriteLn(r.V); end."#
        ),
        &["8"]
    );
}

#[test]
fn class_helper_on_custom_class() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBox=class public Value:Integer; end; TBoxHelper=class helper for TBox procedure Print; end; procedure TBoxHelper.Print; begin WriteLn(Value); end; var b:TBox; begin b:=TBox.Create; b.Value:=77; b.Print; b.Free; end."#
        ),
        &["77"]
    );
}

#[test]
fn enum_helper_name_from_ord() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=(A,B,C); TDHelper=record helper for TD function AsInt:Integer; end; function TDHelper.AsInt:Integer; begin Result:=Ord(Self); end; var d:TD; begin d:=B; WriteLn(d.AsInt); end."#
        ),
        &["1"]
    );
}

#[test]
fn set_helper_contains_all_bits() {
    assert_eq!(
        run_pascal(
            r#"program T; type TSetHelper=record helper for set of Byte function Count:Integer; end; function TSetHelper.Count:Integer; var b:Byte; begin Result:=0; for b in Self do Inc(Result); end; var s:set of Byte; begin s:=[1,2,3]; WriteLn(s.Count); end."#
        ),
        &["3"]
    );
}

#[test]
fn pointer_helper_is_nil() {
    assert_eq!(
        run_pascal(
            r#"program T; type PInt=^Integer; PIntHelper=record helper for PInt function IsNil:Boolean; end; function PIntHelper.IsNil:Boolean; begin Result:=Self=nil; end; var p:PInt; begin WriteLn(p.IsNil); end."#
        ),
        &["true"]
    );
}

#[test]
fn string_helper_trim_spaces_edges() {
    assert_eq!(
        run_pascal(
            r#"program T; type TStrHelper=record helper for String function Trimmed:string; end; function TStrHelper.Trimmed:string; begin Result:=Trim(Self); end; var s:string; begin s:='  hi  '; WriteLn(s.Trimmed); end."#
        ),
        &["hi"]
    );
}

#[test]
fn integer_helper_clamp_range() {
    assert_eq!(
        run_pascal(
            r#"program T; type TIntHelper=record helper for Integer function Clamp(lo,hi:Integer):Integer; end; function TIntHelper.Clamp(lo,hi:Integer):Integer; begin if Self<lo then Result:=lo else if Self>hi then Result:=hi else Result:=Self; end; var n:Integer; begin n:=15; WriteLn(n.Clamp(0,10)); end."#
        ),
        &["10"]
    );
}

#[test]
fn record_helper_compare_equal() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record A,B:Integer; end; TRHelper=record helper for TR function Equals(other:TR):Boolean; end; function TRHelper.Equals(other:TR):Boolean; begin Result:=(A=other.A) and (B=other.B); end; var x,y:TR; begin x.A:=1; x.B:=2; y:=x; WriteLn(x.Equals(y)); end."#
        ),
        &["true"]
    );
}
