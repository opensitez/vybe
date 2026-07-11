/// Class/static methods, class vars, and instance interplay.
use super::helpers::run_pascal;

#[test]
fn class_var_shared_counter_increment() {
    assert_eq!(
        run_pascal(
            r#"program T; type TCounter=class class var N:Integer; class procedure IncN; end; class var TCounter.N:Integer; class procedure TCounter.IncN; begin Inc(N); end; begin TCounter.N:=0; TCounter.IncN; TCounter.IncN; WriteLn(TCounter.N); end."#
        ),
        &["2"]
    );
}

#[test]
fn class_function_factorial_static() {
    assert_eq!(
        run_pascal(
            r#"program T; type TMath=class class function Fact(n:Integer):Integer; end; class function TMath.Fact(n:Integer):Integer; begin if n<=1 then Result:=1 else Result:=n*Fact(n-1); end; begin WriteLn(TMath.Fact(5)); end."#
        ),
        &["120"]
    );
}

#[test]
fn class_procedure_reset_class_var() {
    assert_eq!(
        run_pascal(
            r#"program T; type TStore=class class var V:Integer; class procedure Reset; end; class var TStore.V:Integer; class procedure TStore.Reset; begin V:=0; end; begin TStore.V:=9; TStore.Reset; WriteLn(TStore.V); end."#
        ),
        &["0"]
    );
}

#[test]
fn instance_method_uses_class_var() {
    assert_eq!(
        run_pascal(
            r#"program T; type TApp=class class var Seq:Integer; procedure Next; end; class var TApp.Seq:Integer; procedure TApp.Next; begin Inc(Seq); WriteLn(Seq); end; var a:TApp; begin TApp.Seq:=0; a:=TApp.Create; a.Next; a.Next; end."#
        ),
        &["1", "2"]
    );
}

#[test]
fn class_function_string_repeat() {
    assert_eq!(
        run_pascal(
            r#"program T; type TStr=class class function Dup(s:string; n:Integer):string; end; class function TStr.Dup(s:string; n:Integer):string; var i:Integer; begin Result:=''; for i:=1 to n do Result:=Result+s; end; begin WriteLn(TStr.Dup('o',3)); end."#
        ),
        &["ooo"]
    );
}

#[test]
fn class_const_used_in_method() {
    assert_eq!(
        run_pascal(
            r#"program T; type TLimits=class public const Max=10; class function Cap(n:Integer):Integer; end; class function TLimits.Cap(n:Integer):Integer; begin if n>Max then Result:=Max else Result:=n; end; begin WriteLn(TLimits.Cap(15)); end."#
        ),
        &["10"]
    );
}

#[test]
fn two_instances_share_class_var() {
    assert_eq!(
        run_pascal(
            r#"program T; type TGlobal=class class var Hits:Integer; procedure Ping; end; class var TGlobal.Hits:Integer; procedure TGlobal.Ping; begin Inc(Hits); end; var a,b:TGlobal; begin TGlobal.Hits:=0; a:=TGlobal.Create; b:=TGlobal.Create; a.Ping; b.Ping; WriteLn(TGlobal.Hits); end."#
        ),
        &["2"]
    );
}

#[test]
fn class_function_min_pair() {
    assert_eq!(
        run_pascal(
            r#"program T; type TUtil=class class function Min(a,b:Integer):Integer; end; class function TUtil.Min(a,b:Integer):Integer; begin if a<b then Result:=a else Result:=b; end; begin WriteLn(TUtil.Min(7,3)); end."#
        ),
        &["3"]
    );
}

#[test]
fn class_procedure_log_line() {
    assert_eq!(
        run_pascal(
            r#"program T; type TLog=class class procedure Line(const s:string); end; class procedure TLog.Line(const s:string); begin WriteLn('L:'+s); end; begin TLog.Line('ok'); end."#
        ),
        &["L:ok"]
    );
}

#[test]
fn class_var_initialized_nonzero() {
    assert_eq!(
        run_pascal(
            r#"program T; type TSeed=class class var Base:Integer; end; class var TSeed.Base:Integer; begin TSeed.Base:=100; WriteLn(TSeed.Base); end."#
        ),
        &["100"]
    );
}

#[test]
fn instance_calls_class_function() {
    assert_eq!(
        run_pascal(
            r#"program T; type TWrap=class function Double(n:Integer):Integer; end; function TWrap.Double(n:Integer):Integer; begin Result:=n*2; end; var w:TWrap; begin w:=TWrap.Create; WriteLn(w.Double(6)); end."#
        ),
        &["12"]
    );
}

#[test]
fn class_function_is_even() {
    assert_eq!(
        run_pascal(
            r#"program T; type TNum=class class function Even(n:Integer):Boolean; end; class function TNum.Even(n:Integer):Boolean; begin Result:=(n mod 2)=0; end; begin WriteLn(TNum.Even(8)); WriteLn(TNum.Even(7)); end."#
        ),
        &["true", "false"]
    );
}

#[test]
fn class_var_accumulate_from_method() {
    assert_eq!(
        run_pascal(
            r#"program T; type TSum=class class var Total:Integer; procedure Add(v:Integer); end; class var TSum.Total:Integer; procedure TSum.Add(v:Integer); begin Total:=Total+v; end; var s:TSum; begin TSum.Total:=0; s:=TSum.Create; s.Add(3); s.Add(4); WriteLn(TSum.Total); end."#
        ),
        &["7"]
    );
}

#[test]
fn class_function_abs_value() {
    assert_eq!(
        run_pascal(
            r#"program T; type TAbs=class class function Val(n:Integer):Integer; end; class function TAbs.Val(n:Integer):Integer; begin if n<0 then Result:=-n else Result:=n; end; begin WriteLn(TAbs.Val(-9)); end."#
        ),
        &["9"]
    );
}

#[test]
fn class_procedure_swap_class_fields() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPair=class class var A,B:Integer; class procedure Swap; end; class var TPair.A:Integer; class var TPair.B:Integer; class procedure TPair.Swap; var t:Integer; begin t:=A; A:=B; B:=t; end; begin TPair.A:=1; TPair.B:=2; TPair.Swap; WriteLn(TPair.A); WriteLn(TPair.B); end."#
        ),
        &["2", "1"]
    );
}

#[test]
fn class_function_concat_tags() {
    assert_eq!(
        run_pascal(
            r#"program T; type TTag=class class function Join(a,b:string):string; end; class function TTag.Join(a,b:string):string; begin Result:='<'+a+'/'+b+'>'; end; begin WriteLn(TTag.Join('a','b')); end."#
        ),
        &["<a/b>"]
    );
}

#[test]
fn class_method_chain_on_instance() {
    assert_eq!(
        run_pascal(
            r#"program T; type TAcc=class Value:Integer; procedure Add(n:Integer); function Get:Integer; end; procedure TAcc.Add(n:Integer); begin Value:=Value+n; end; function TAcc.Get:Integer; begin Result:=Value; end; var a:TAcc; begin a:=TAcc.Create; a.Value:=1; a.Add(2); WriteLn(a.Get); end."#
        ),
        &["3"]
    );
}

#[test]
fn class_function_power_small() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPow=class class function P(n,e:Integer):Integer; end; class function TPow.P(n,e:Integer):Integer; var i:Integer; begin Result:=1; for i:=1 to e do Result:=Result*n; end; begin WriteLn(TPow.P(2,4)); end."#
        ),
        &["16"]
    );
}

#[test]
fn class_var_name_prefix() {
    assert_eq!(
        run_pascal(
            r#"program T; type TId=class class var Next:Integer; class function Gen:string; end; class var TId.Next:Integer; class function TId.Gen:string; begin Result:='id'+IntToStr(Next); Inc(Next); end; begin TId.Next:=1; WriteLn(TId.Gen); WriteLn(TId.Gen); end."#
        ),
        &["id1", "id2"]
    );
}

#[test]
fn class_procedure_clear_flag() {
    assert_eq!(
        run_pascal(
            r#"program T; type TFlag=class class var On:Boolean; class procedure Off; end; class var TFlag.On:Boolean; class procedure TFlag.Off; begin On:=false; end; begin TFlag.On:=true; TFlag.Off; WriteLn(TFlag.On); end."#
        ),
        &["false"]
    );
}

#[test]
fn class_function_gcd_style() {
    assert_eq!(
        run_pascal(
            r#"program T; type TGcd=class class function G(a,b:Integer):Integer; end; class function TGcd.G(a,b:Integer):Integer; begin if b=0 then Result:=a else Result:=G(b,a mod b); end; begin WriteLn(TGcd.G(48,18)); end."#
        ),
        &["6"]
    );
}

#[test]
fn instance_field_vs_class_var() {
    assert_eq!(
        run_pascal(
            r#"program T; type TMix=class class var Shared:Integer; Local:Integer; end; class var TMix.Shared:Integer; var a,b:TMix; begin TMix.Shared:=0; a:=TMix.Create; b:=TMix.Create; a.Local:=1; b.Local:=2; Inc(TMix.Shared); WriteLn(a.Local); WriteLn(TMix.Shared); end."#
        ),
        &["1", "1"]
    );
}

#[test]
fn class_function_return_bool_compare() {
    assert_eq!(
        run_pascal(
            r#"program T; type TCmp=class class function Eq(a,b:Integer):Boolean; end; class function TCmp.Eq(a,b:Integer):Boolean; begin Result:=a=b; end; begin WriteLn(TCmp.Eq(3,3)); WriteLn(TCmp.Eq(1,2)); end."#
        ),
        &["true", "false"]
    );
}

#[test]
fn class_procedure_print_twice() {
    assert_eq!(
        run_pascal(
            r#"program T; type TOut=class class procedure Twice(const s:string); end; class procedure TOut.Twice(const s:string); begin WriteLn(s); WriteLn(s); end; begin TOut.Twice('x'); end."#
        ),
        &["x", "x"]
    );
}

#[test]
fn class_function_clamp_range() {
    assert_eq!(
        run_pascal(
            r#"program T; type TClamp=class class function InRange(v,lo,hi:Integer):Integer; end; class function TClamp.InRange(v,lo,hi:Integer):Integer; begin if v<lo then Result:=lo else if v>hi then Result:=hi else Result:=v; end; begin WriteLn(TClamp.InRange(15,0,10)); end."#
        ),
        &["10"]
    );
}

#[test]
fn class_var_set_from_static_method() {
    assert_eq!(
        run_pascal(
            r#"program T; type TCfg=class class var Port:Integer; class procedure SetPort(p:Integer); end; class var TCfg.Port:Integer; class procedure TCfg.SetPort(p:Integer); begin Port:=p; end; begin TCfg.SetPort(8080); WriteLn(TCfg.Port); end."#
        ),
        &["8080"]
    );
}

#[test]
fn class_function_negate_int() {
    assert_eq!(
        run_pascal(
            r#"program T; type TNeg=class class function Flip(n:Integer):Integer; end; class function TNeg.Flip(n:Integer):Integer; begin Result:=-n; end; begin WriteLn(TNeg.Flip(11)); end."#
        ),
        &["-11"]
    );
}

#[test]
fn class_method_read_instance_field() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBox=class W,H:Integer; function Area:Integer; end; function TBox.Area:Integer; begin Result:=W*H; end; var b:TBox; begin b:=TBox.Create; b.W:=4; b.H:=5; WriteLn(b.Area); end."#
        ),
        &["20"]
    );
}

#[test]
fn class_function_char_to_upper() {
    assert_eq!(
        run_pascal(
            r#"program T; type TCharUtil=class class function Up(c:Char):Char; end; class function TCharUtil.Up(c:Char):Char; begin Result:=UpCase(c); end; begin WriteLn(TCharUtil.Up('g')); end."#
        ),
        &["G"]
    );
}

#[test]
fn class_var_decrement_procedure() {
    assert_eq!(
        run_pascal(
            r#"program T; type TDec=class class var N:Integer; class procedure Step; end; class var TDec.N:Integer; class procedure TDec.Step; begin Dec(N); end; begin TDec.N:=5; TDec.Step; TDec.Step; WriteLn(TDec.N); end."#
        ),
        &["3"]
    );
}

#[test]
fn class_function_sum_three() {
    assert_eq!(
        run_pascal(
            r#"program T; type TAdd=class class function Sum(a,b,c:Integer):Integer; end; class function TAdd.Sum(a,b,c:Integer):Integer; begin Result:=a+b+c; end; begin WriteLn(TAdd.Sum(1,2,3)); end."#
        ),
        &["6"]
    );
}

#[test]
fn class_procedure_toggle_class_bool() {
    assert_eq!(
        run_pascal(
            r#"program T; type TToggle=class class var On:Boolean; class procedure Flip; end; class var TToggle.On:Boolean; class procedure TToggle.Flip; begin On:=not On; end; begin TToggle.On:=false; TToggle.Flip; WriteLn(TToggle.On); end."#
        ),
        &["true"]
    );
}

#[test]
fn class_function_length_of_string() {
    assert_eq!(
        run_pascal(
            r#"program T; type TLen=class class function OfStr(const s:string):Integer; end; class function TLen.OfStr(const s:string):Integer; begin Result:=Length(s); end; begin WriteLn(TLen.OfStr('abcd')); end."#
        ),
        &["4"]
    );
}

#[test]
fn instance_constructor_sets_field() {
    assert_eq!(
        run_pascal(
            r#"program T; type TItem=class constructor Create(v:Integer); Value:Integer; end; constructor TItem.Create(v:Integer); begin Value:=v; end; var i:TItem; begin i:=TItem.Create(42); WriteLn(i.Value); end."#
        ),
        &["42"]
    );
}

#[test]
fn class_function_modulo() {
    assert_eq!(
        run_pascal(
            r#"program T; type TMod=class class function Rem(a,b:Integer):Integer; end; class function TMod.Rem(a,b:Integer):Integer; begin Result:=a mod b; end; begin WriteLn(TMod.Rem(10,3)); end."#
        ),
        &["1"]
    );
}

#[test]
fn class_var_max_tracking() {
    assert_eq!(
        run_pascal(
            r#"program T; type TMax=class class var Best:Integer; class procedure Consider(v:Integer); end; class var TMax.Best:Integer; class procedure TMax.Consider(v:Integer); begin if v>Best then Best:=v; end; begin TMax.Best:=0; TMax.Consider(3); TMax.Consider(7); TMax.Consider(5); WriteLn(TMax.Best); end."#
        ),
        &["7"]
    );
}

#[test]
fn class_function_is_positive() {
    assert_eq!(
        run_pascal(
            r#"program T; type TSign=class class function Pos(n:Integer):Boolean; end; class function TSign.Pos(n:Integer):Boolean; begin Result:=n>0; end; begin WriteLn(TSign.Pos(1)); WriteLn(TSign.Pos(0)); end."#
        ),
        &["true", "false"]
    );
}

#[test]
fn class_procedure_emit_numbered() {
    assert_eq!(
        run_pascal(
            r#"program T; type TSeq=class class var I:Integer; class procedure Tick; end; class var TSeq.I:Integer; class procedure TSeq.Tick; begin Inc(I); WriteLn(I); end; begin TSeq.I:=0; TSeq.Tick; TSeq.Tick; end."#
        ),
        &["1", "2"]
    );
}

#[test]
fn class_function_average_two() {
    assert_eq!(
        run_pascal(
            r#"program T; type TAvg=class class function Mean(a,b:Integer):Integer; end; class function TAvg.Mean(a,b:Integer):Integer; begin Result:=(a+b) div 2; end; begin WriteLn(TAvg.Mean(5,9)); end."#
        ),
        &["7"]
    );
}

#[test]
fn class_method_mutate_own_field() {
    assert_eq!(
        run_pascal(
            r#"program T; type TInc=class N:Integer; procedure Bump; end; procedure TInc.Bump; begin Inc(N); end; var x:TInc; begin x:=TInc.Create; x.N:=10; x.Bump; WriteLn(x.N); end."#
        ),
        &["11"]
    );
}

#[test]
fn class_function_first_char() {
    assert_eq!(
        run_pascal(
            r#"program T; type THead=class class function First(const s:string):Char; end; class function THead.First(const s:string):Char; begin Result:=s[1]; end; begin WriteLn(THead.First('zoo')); end."#
        ),
        &["z"]
    );
}

#[test]
fn class_var_string_buffer_append() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBuf=class class var S:string; class procedure Add(const t:string); end; class var TBuf.S:string; class procedure TBuf.Add(const t:string); begin S:=S+t; end; begin TBuf.S:=''; TBuf.Add('a'); TBuf.Add('b'); WriteLn(TBuf.S); end."#
        ),
        &["ab"]
    );
}
