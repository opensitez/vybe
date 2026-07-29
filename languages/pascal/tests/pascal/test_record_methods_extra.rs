/// Record methods, helpers, and small utilities on records.
use super::helpers::run_pascal;

#[test]
fn record_method_sum_fields() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPair=record A,B:Integer; function Sum:Integer; end; function TPair.Sum:Integer; begin Result:=A+B; end; var p:TPair; begin p.A:=6; p.B:=7; WriteLn(p.Sum); end."#
        ),
        &["13"]
    );
}

#[test]
fn record_procedure_mutate_field() {
    assert_eq!(
        run_pascal(
            r#"program T; type TAcc=record N:Integer; procedure IncN; end; procedure TAcc.IncN; begin Inc(N); end; var a:TAcc; begin a.N:=0; a.IncN; a.IncN; WriteLn(a.N); end."#
        ),
        &["2"]
    );
}

#[test]
fn record_helper_distance_squared() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPt=record X,Y:Integer; end; TPtHelper=record helper for TPt function LenSq:Integer; end; function TPtHelper.LenSq:Integer; begin Result:=X*X+Y*Y; end; var p:TPt; begin p.X:=3; p.Y:=4; WriteLn(p.LenSq); end."#
        ),
        &["25"]
    );
}

#[test]
fn record_method_product() {
    assert_eq!(
        run_pascal(
            r#"program T; type TRect=record W,H:Integer; function Area:Integer; end; function TRect.Area:Integer; begin Result:=W*H; end; var r:TRect; begin r.W:=5; r.H:=8; WriteLn(r.Area); end."#
        ),
        &["40"]
    );
}

#[test]
fn record_function_factory() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPt=record X,Y:Integer; end; function MakePt(x,y:Integer):TPt; begin Result.X:=x; Result.Y:=y; end; var p:TPt; begin p:=MakePt(2,3); WriteLn(p.X+p.Y); end."#
        ),
        &["5"]
    );
}

#[test]
fn record_method_is_origin() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPt=record X,Y:Integer; function IsZero:Boolean; end; function TPt.IsZero:Boolean; begin Result:=(X=0) and (Y=0); end; var p:TPt; begin p.X:=0; p.Y:=0; WriteLn(p.IsZero); end."#
        ),
        &["True"]
    );
}

#[test]
fn record_helper_swap_xy() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPt=record X,Y:Integer; end; TPtHelper=record helper for TPt procedure Swap; end; procedure TPtHelper.Swap; var t:Integer; begin t:=X; X:=Y; Y:=t; end; var p:TPt; begin p.X:=1; p.Y:=9; p.Swap; WriteLn(p.X); WriteLn(p.Y); end."#
        ),
        &["9", "1"]
    );
}

#[test]
fn record_method_scale_by_int() {
    assert_eq!(
        run_pascal(
            r#"program T; type TVec=record X,Y:Integer; procedure Scale(k:Integer); end; procedure TVec.Scale(k:Integer); begin X:=X*k; Y:=Y*k; end; var v:TVec; begin v.X:=2; v.Y:=3; v.Scale(4); WriteLn(v.X); WriteLn(v.Y); end."#
        ),
        &["8", "12"]
    );
}

#[test]
fn record_nested_method_delegate() {
    assert_eq!(
        run_pascal(
            r#"program T; type TInner=record V:Integer; function Get:Integer; end; function TInner.Get:Integer; begin Result:=V; end; type TOuter=record Inner:TInner; function Total:Integer; end; function TOuter.Total:Integer; begin Result:=Inner.Get; end; var o:TOuter; begin o.Inner.V:=11; WriteLn(o.Total); end."#
        ),
        &["11"]
    );
}

#[test]
fn record_helper_string_from_point() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPt=record X,Y:Integer; end; TPtHelper=record helper for TPt function AsText:string; end; function TPtHelper.AsText:string; begin Result:=IntToStr(X)+','+IntToStr(Y); end; var p:TPt; begin p.X:=2; p.Y:=5; WriteLn(p.AsText); end."#
        ),
        &["2,5"]
    );
}

#[test]
fn record_method_reset_fields() {
    assert_eq!(
        run_pascal(
            r#"program T; type TReset=record A,B:Integer; procedure Clear; end; procedure TReset.Clear; begin A:=0; B:=0; end; var r:TReset; begin r.A:=5; r.B:=6; r.Clear; WriteLn(r.A+r.B); end."#
        ),
        &["0"]
    );
}

#[test]
fn record_function_max_component() {
    assert_eq!(
        run_pascal(
            r#"program T; type TRng=record Lo,Hi:Integer; function Span:Integer; end; function TRng.Span:Integer; begin Result:=Hi-Lo; end; var r:TRng; begin r.Lo:=3; r.Hi:=10; WriteLn(r.Span); end."#
        ),
        &["7"]
    );
}

#[test]
fn record_helper_compare_x() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPt=record X,Y:Integer; end; TPtHelper=record helper for TPt function SameX(const o:TPt):Boolean; end; function TPtHelper.SameX(const o:TPt):Boolean; begin Result:=X=o.X; end; var a,b:TPt; begin a.X:=1; a.Y:=2; b.X:=1; b.Y:=9; WriteLn(a.SameX(b)); end."#
        ),
        &["True"]
    );
}

#[test]
fn record_method_toggle_flag() {
    assert_eq!(
        run_pascal(
            r#"program T; type TFlag=record On:Boolean; procedure Flip; end; procedure TFlag.Flip; begin On:=not On; end; var f:TFlag; begin f.On:=false; f.Flip; WriteLn(f.On); end."#
        ),
        &["True"]
    );
}

#[test]
fn record_var_param_bump() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPt=record X:Integer; end; procedure Bump(var p:TPt); begin Inc(p.X); end; var p:TPt; begin p.X:=4; Bump(p); WriteLn(p.X); end."#
        ),
        &["5"]
    );
}

#[test]
fn record_method_char_field_upper() {
    assert_eq!(
        run_pascal(
            r#"program T; type TCh=record C:Char; function Upper:Char; end; function TCh.Upper:Char; begin Result:=UpCase(C); end; var c:TCh; begin c.C:='m'; WriteLn(c.Upper); end."#
        ),
        &["M"]
    );
}

#[test]
fn record_helper_add_to_x() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPt=record X,Y:Integer; end; TPtHelper=record helper for TPt procedure AddX(d:Integer); end; procedure TPtHelper.AddX(d:Integer); begin X:=X+d; end; var p:TPt; begin p.X:=1; p.AddX(4); WriteLn(p.X); end."#
        ),
        &["5"]
    );
}

#[test]
fn record_method_string_field_len() {
    assert_eq!(
        run_pascal(
            r#"program T; type TName=record S:string; function Len:Integer; end; function TName.Len:Integer; begin Result:=Length(S); end; var n:TName; begin n.S:='abc'; WriteLn(n.Len); end."#
        ),
        &["3"]
    );
}

#[test]
fn record_nested_clear_inner() {
    assert_eq!(
        run_pascal(
            r#"program T; type TInner=record V:Integer; procedure Zero; end; procedure TInner.Zero; begin V:=0; end; type TWrap=record Inner:TInner; procedure Reset; end; procedure TWrap.Reset; begin Inner.Zero; end; var w:TWrap; begin w.Inner.V:=9; w.Reset; WriteLn(w.Inner.V); end."#
        ),
        &["0"]
    );
}

#[test]
fn record_function_copy_value() {
    assert_eq!(
        run_pascal(
            r#"program T; type TVal=record N:Integer; end; function Dup(v:TVal):TVal; begin Result:=v; Result.N:=Result.N+1; end; var a,b:TVal; begin a.N:=5; b:=Dup(a); WriteLn(b.N); end."#
        ),
        &["6"]
    );
}

#[test]
fn record_method_average_two_fields() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPair=record A,B:Integer; function Avg:Integer; end; function TPair.Avg:Integer; begin Result:=(A+B) div 2; end; var p:TPair; begin p.A:=4; p.B:=8; WriteLn(p.Avg); end."#
        ),
        &["6"]
    );
}

#[test]
fn record_helper_is_positive_x() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPt=record X,Y:Integer; end; TPtHelper=record helper for TPt function XPos:Boolean; end; function TPtHelper.XPos:Boolean; begin Result:=X>0; end; var p:TPt; begin p.X:=3; WriteLn(p.XPos); end."#
        ),
        &["True"]
    );
}

#[test]
fn record_method_increment_both() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPair=record A,B:Integer; procedure IncBoth; end; procedure TPair.IncBoth; begin Inc(A); Inc(B); end; var p:TPair; begin p.A:=1; p.B:=2; p.IncBoth; WriteLn(p.A); WriteLn(p.B); end."#
        ),
        &["2", "3"]
    );
}

#[test]
fn record_method_min_field() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPair=record A,B:Integer; function MinF:Integer; end; function TPair.MinF:Integer; begin if A<B then Result:=A else Result:=B; end; var p:TPair; begin p.A:=7; p.B:=3; WriteLn(p.MinF); end."#
        ),
        &["3"]
    );
}

#[test]
fn record_helper_double_values() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPair=record A,B:Integer; end; TPairHelper=record helper for TPair procedure Double; end; procedure TPairHelper.Double; begin A:=A*2; B:=B*2; end; var p:TPair; begin p.A:=2; p.B:=3; p.Double; WriteLn(p.A+p.B); end."#
        ),
        &["10"]
    );
}

#[test]
fn record_method_set_string() {
    assert_eq!(
        run_pascal(
            r#"program T; type TMsg=record Text:string; procedure SetText(const s:string); end; procedure TMsg.SetText(const s:string); begin Text:=s; end; var m:TMsg; begin m.SetText('hi'); WriteLn(m.Text); end."#
        ),
        &["hi"]
    );
}

#[test]
fn record_function_add_points() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPt=record X,Y:Integer; end; function Add(a,b:TPt):TPt; begin Result.X:=a.X+b.X; Result.Y:=a.Y+b.Y; end; var u,v,w:TPt; begin u.X:=1; u.Y:=2; v.X:=3; v.Y:=4; w:=Add(u,v); WriteLn(w.X); WriteLn(w.Y); end."#
        ),
        &["4", "6"]
    );
}

#[test]
fn record_method_abs_x() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPt=record X:Integer; function AbsX:Integer; end; function TPt.AbsX:Integer; begin if X<0 then Result:=-X else Result:=X; end; var p:TPt; begin p.X:=-8; WriteLn(p.AbsX); end."#
        ),
        &["8"]
    );
}

#[test]
fn record_helper_equal_coords() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPt=record X,Y:Integer; end; TPtHelper=record helper for TPt function Equal(const o:TPt):Boolean; end; function TPtHelper.Equal(const o:TPt):Boolean; begin Result:=(X=o.X) and (Y=o.Y); end; var a,b:TPt; begin a.X:=1; a.Y:=2; b:=a; WriteLn(a.Equal(b)); end."#
        ),
        &["True"]
    );
}

#[test]
fn record_method_shift_left() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPt=record X,Y:Integer; procedure Shift(dx,dy:Integer); end; procedure TPt.Shift(dx,dy:Integer); begin X:=X+dx; Y:=Y+dy; end; var p:TPt; begin p.X:=0; p.Y:=0; p.Shift(2,3); WriteLn(p.X); WriteLn(p.Y); end."#
        ),
        &["2", "3"]
    );
}

#[test]
fn record_method_count_bits_simple() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBits=record Mask:Integer; function Pop:Integer; end; function TBits.Pop:Integer; var c,m:Integer; begin c:=0; m:=Mask; while m>0 do begin if (m mod 2)=1 then Inc(c); m:=m div 2; end; Result:=c; end; var b:TBits; begin b.Mask:=5; WriteLn(b.Pop); end."#
        ),
        &["2"]
    );
}

#[test]
fn record_helper_init_defaults() {
    assert_eq!(
        run_pascal(
            r#"program T; type TCfg=record A,B:Integer; end; TCfgHelper=record helper for TCfg procedure Init; end; procedure TCfgHelper.Init; begin A:=1; B:=2; end; var c:TCfg; begin c.Init; WriteLn(c.A+c.B); end."#
        ),
        &["3"]
    );
}

#[test]
fn record_method_compare_y() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPt=record X,Y:Integer; function YGreater(const o:TPt):Boolean; end; function TPt.YGreater(const o:TPt):Boolean; begin Result:=Y>o.Y; end; var a,b:TPt; begin a.Y:=5; b.Y:=2; WriteLn(a.YGreater(b)); end."#
        ),
        &["True"]
    );
}

#[test]
fn record_method_concat_names() {
    assert_eq!(
        run_pascal(
            r#"program T; type TFull=record First,Last:string; function Full:string; end; function TFull.Full:string; begin Result:=First+' '+Last; end; var f:TFull; begin f.First:='Ann'; f.Last:='Lee'; WriteLn(f.Full); end."#
        ),
        &["Ann Lee"]
    );
}

#[test]
fn record_function_midpoint() {
    assert_eq!(
        run_pascal(
            r#"program T; type TRng=record Lo,Hi:Integer; function Mid:Integer; end; function TRng.Mid:Integer; begin Result:=(Lo+Hi) div 2; end; var r:TRng; begin r.Lo:=2; r.Hi:=8; WriteLn(r.Mid); end."#
        ),
        &["5"]
    );
}

#[test]
fn record_helper_has_zero() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPair=record A,B:Integer; end; TPairHelper=record helper for TPair function HasZero:Boolean; end; function TPairHelper.HasZero:Boolean; begin Result:=(A=0) or (B=0); end; var p:TPair; begin p.A:=0; p.B:=3; WriteLn(p.HasZero); end."#
        ),
        &["True"]
    );
}

#[test]
fn record_method_divide_safe() {
    assert_eq!(
        run_pascal(
            r#"program T; type TFrac=record Num,Den:Integer; function Value:Integer; end; function TFrac.Value:Integer; begin if Den=0 then Result:=0 else Result:=Num div Den; end; var f:TFrac; begin f.Num:=9; f.Den:=3; WriteLn(f.Value); end."#
        ),
        &["3"]
    );
}

#[test]
fn record_method_char_to_ord() {
    assert_eq!(
        run_pascal(
            r#"program T; type TCh=record C:Char; function Code:Integer; end; function TCh.Code:Integer; begin Result:=Ord(C); end; var c:TCh; begin c.C:='A'; WriteLn(c.Code); end."#
        ),
        &["65"]
    );
}

#[test]
fn record_helper_magnitude_manhattan() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPt=record X,Y:Integer; end; TPtHelper=record helper for TPt function Manh:Integer; end; function TPtHelper.Manh:Integer; begin Result:=Abs(X)+Abs(Y); end; var p:TPt; begin p.X:=-2; p.Y:=3; WriteLn(p.Manh); end."#
        ),
        &["5"]
    );
}

#[test]
fn record_method_clear_string() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBuf=record S:string; procedure Clear; end; procedure TBuf.Clear; begin S:=''; end; var b:TBuf; begin b.S:='x'; b.Clear; WriteLn(Length(b.S)); end."#
        ),
        &["0"]
    );
}

#[test]
fn record_method_sum_three_fields() {
    assert_eq!(
        run_pascal(
            r#"program T; type TTrip=record A,B,C:Integer; function Sum:Integer; end; function TTrip.Sum:Integer; begin Result:=A+B+C; end; var t:TTrip; begin t.A:=1; t.B:=2; t.C:=3; WriteLn(t.Sum); end."#
        ),
        &["6"]
    );
}

#[test]
fn record_helper_negate_x() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPt=record X,Y:Integer; end; TPtHelper=record helper for TPt procedure NegX; end; procedure TPtHelper.NegX; begin X:=-X; end; var p:TPt; begin p.X:=4; p.NegX; WriteLn(p.X); end."#
        ),
        &["-4"]
    );
}
