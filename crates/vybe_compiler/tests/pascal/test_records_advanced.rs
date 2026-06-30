/// Case records, nested records, variant parts — beyond test_records_types.rs.
use super::helpers::run_pascal;

#[test]
fn case_record_tag_zero_branch_fields() {
    assert_eq!(
        run_pascal(r#"program T; type TVal=record case K:Integer of 0:(I:Integer); 1:(S:string); end; var v:TVal; begin v.K:=0; v.I:=42; WriteLn(v.I); end."#),
        &["42"]
    );
}

#[test]
fn case_record_tag_one_string_branch() {
    assert_eq!(
        run_pascal(r#"program T; type TVal=record case K:Integer of 0:(I:Integer); 1:(S:string); end; var v:TVal; begin v.K:=1; v.S:='hi'; WriteLn(v.S); end."#),
        &["hi"]
    );
}

#[test]
fn case_record_boolean_discriminant() {
    assert_eq!(
        run_pascal(r#"program T; type TNum=record case IsReal:Boolean of false:(N:Integer); true:(R:Double); end; var x:TNum; begin x.IsReal:=false; x.N:=7; WriteLn(x.N); end."#),
        &["7"]
    );
}

#[test]
fn case_record_enum_discriminant() {
    assert_eq!(
        run_pascal(r#"program T; type TKind=(IntK,StrK); type TBox=record case Kind:TKind of IntK:(V:Integer); StrK:(T:string); end; var b:TBox; begin b.Kind:=StrK; b.T:='ok'; WriteLn(b.T); end."#),
        &["ok"]
    );
}

#[test]
fn case_record_with_fixed_prefix_field() {
    assert_eq!(
        run_pascal(r#"program T; type TMsg=record Id:Integer; case Kind:Integer of 1:(Text:string); 2:(Code:Integer); end; var m:TMsg; begin m.Id:=9; m.Kind:=2; m.Code:=404; WriteLn(m.Code); end."#),
        &["404"]
    );
}

#[test]
fn nested_record_two_levels() {
    assert_eq!(
        run_pascal(r#"program T; type TInner=record V:Integer; end; type TOuter=record Inner:TInner; end; var o:TOuter; begin o.Inner.V:=55; WriteLn(o.Inner.V); end."#),
        &["55"]
    );
}

#[test]
fn nested_record_assignment_copies_inner() {
    assert_eq!(
        run_pascal(r#"program T; type TInner=record N:Integer; end; type TOuter=record A,B:TInner; end; var o:TOuter; begin o.A.N:=3; o.B:=o.A; o.B.N:=9; WriteLn(o.A.N); end."#),
        &["3"]
    );
}

#[test]
fn record_with_method_sum_fields() {
    assert_eq!(
        run_pascal(r#"program T; type TPair=record A,B:Integer; function Sum:Integer; end; function TPair.Sum:Integer; begin Result:=A+B; end; var p:TPair; begin p.A:=4; p.B:=5; WriteLn(p.Sum); end."#),
        &["9"]
    );
}

#[test]
fn record_method_mutates_field() {
    assert_eq!(
        run_pascal(r#"program T; type TAcc=record N:Integer; procedure IncN; end; procedure TAcc.IncN; begin N:=N+1; end; var a:TAcc; begin a.N:=1; a.IncN; WriteLn(a.N); end."#),
        &["2"]
    );
}

#[test]
fn case_record_three_way_tag() {
    assert_eq!(
        run_pascal(r#"program T; type TShape=record case Tag:Integer of 0:(W,H:Integer); 1:(R:Integer); 2:(Side:Integer); end; var s:TShape; begin s.Tag:=2; s.Side:=6; WriteLn(s.Side); end."#),
        &["6"]
    );
}

#[test]
fn case_record_switch_on_tag() {
    assert_eq!(
        run_pascal(r#"program T; type TShape=record case Tag:Integer of 0:(R:Integer); 1:(W,H:Integer); end; var s:TShape; begin s.Tag:=1; s.W:=3; s.H:=4; if s.Tag=1 then WriteLn(s.W*s.H) else WriteLn(0); end."#),
        &["12"]
    );
}

#[test]
fn nested_record_three_deep() {
    assert_eq!(
        run_pascal(r#"program T; type TL3=record V:Integer; end; type TL2=record L3:TL3; end; type TL1=record L2:TL2; end; var r:TL1; begin r.L2.L3.V:=21; WriteLn(r.L2.L3.V); end."#),
        &["21"]
    );
}

#[test]
fn record_array_of_records() {
    assert_eq!(
        run_pascal(r#"program T; type TItem=record V:Integer; end; var a:array[0..1] of TItem; begin a[0].V:=1; a[1].V:=2; WriteLn(a[1].V); end."#),
        &["2"]
    );
}

#[test]
fn record_containing_static_array() {
    assert_eq!(
        run_pascal(r#"program T; type TBuf=record Data:array[0..2] of Integer; end; var b:TBuf; begin b.Data[2]:=8; WriteLn(b.Data[2]); end."#),
        &["8"]
    );
}

#[test]
fn record_with_string_field_concat() {
    assert_eq!(
        run_pascal(r#"program T; type TName=record First,Last:string; end; var n:TName; begin n.First:='Ann'; n.Last:='Lee'; WriteLn(n.First+' '+n.Last); end."#),
        &["Ann Lee"]
    );
}

#[test]
fn case_record_char_discriminant() {
    assert_eq!(
        run_pascal(r#"program T; type TTok=record case Kind:Char of 'I':(N:Integer); 'S':(T:string); end; var t:TTok; begin t.Kind:='I'; t.N:=11; WriteLn(t.N); end."#),
        &["11"]
    );
}

#[test]
fn record_equal_fields_independent() {
    assert_eq!(
        run_pascal(r#"program T; type TPt=record X,Y:Integer; end; var a,b:TPt; begin a.X:=1; a.Y:=2; b:=a; b.X:=9; WriteLn(a.X); WriteLn(b.X); end."#),
        &["1", "9"]
    );
}

#[test]
fn nested_record_in_case_variant() {
    assert_eq!(
        run_pascal(r#"program T; type TInner=record V:Integer; end; type TWrap=record case K:Integer of 0:(I:Integer); 1:(Inner:TInner); end; var w:TWrap; begin w.K:=1; w.Inner.V:=77; WriteLn(w.Inner.V); end."#),
        &["77"]
    );
}

#[test]
fn record_procedure_param_field_update() {
    assert_eq!(
        run_pascal(r#"program T; type TPt=record X:Integer; end; procedure Bump(var p:TPt); begin p.X:=p.X+1; end; var p:TPt; begin p.X:=4; Bump(p); WriteLn(p.X); end."#),
        &["5"]
    );
}

#[test]
fn record_function_returns_new() {
    assert_eq!(
        run_pascal(r#"program T; type TPt=record X,Y:Integer; end; function Make(x,y:Integer):TPt; begin Result.X:=x; Result.Y:=y; end; var p:TPt; begin p:=Make(2,3); WriteLn(p.X+p.Y); end."#),
        &["5"]
    );
}

#[test]
fn case_record_overlapping_integer_variant() {
    assert_eq!(
        run_pascal(r#"program T; type TNum=record case T:Integer of 0,1:(Lo,Hi:Byte); 2:(Value:Integer); end; var n:TNum; begin n.T:=2; n.Value:=1000; WriteLn(n.Value); end."#),
        &["1000"]
    );
}

#[test]
fn record_boolean_fields_and() {
    assert_eq!(
        run_pascal(r#"program T; type TFlags=record A,B:Boolean; end; var f:TFlags; begin f.A:=true; f.B:=false; WriteLn(f.A and f.B); end."#),
        &["false"]
    );
}

#[test]
fn record_real_field_round() {
    assert_eq!(
        run_pascal(r#"program T; type TMeas=record V:Double; end; var m:TMeas; begin m.V:=2.6; WriteLn(Round(m.V)); end."#),
        &["3"]
    );
}

#[test]
fn nested_record_sum_via_method() {
    assert_eq!(
        run_pascal(r#"program T; type TInner=record V:Integer; function Get:Integer; end; function TInner.Get:Integer; begin Result:=V; end; type TOuter=record Inner:TInner; function Total:Integer; end; function TOuter.Total:Integer; begin Result:=Inner.Get; end; var o:TOuter; begin o.Inner.V:=6; WriteLn(o.Total); end."#),
        &["6"]
    );
}

#[test]
fn case_record_tag_then_branch_field_read() {
    assert_eq!(
        run_pascal(r#"program T; type TE=record case K:Integer of 1:(A:Integer); 2:(B:Integer); end; var e:TE; begin e.K:=2; e.B:=31; WriteLn(e.B); end."#),
        &["31"]
    );
}

#[test]
fn record_with_set_field() {
    assert_eq!(
        run_pascal(r#"program T; type TDays=set of (Mon,Tue,Wed); type TSched=record Open:TDays; end; var s:TSched; begin s.Open:=[Mon,Wed]; WriteLn(Ord(Mon) in s.Open); WriteLn(Ord(Tue) in s.Open); end."#),
        &["true", "false"]
    );
}

#[test]
fn record_variant_part_with_string_and_int() {
    assert_eq!(
        run_pascal(r#"program T; type TCell=record case Active:Boolean of false:(I:Integer); true:(S:string); end; var c:TCell; begin c.Active:=true; c.S:='text'; WriteLn(c.S); end."#),
        &["text"]
    );
}

#[test]
fn record_pass_by_value_copy_isolated() {
    assert_eq!(
        run_pascal(r#"program T; type TBox=record V:Integer; end; function Double(b:TBox):TBox; begin b.V:=b.V*2; Result:=b; end; var x:TBox; begin x.V:=5; x:=Double(x); WriteLn(x.V); end."#),
        &["10"]
    );
}

#[test]
fn nested_record_reset_inner_only() {
    assert_eq!(
        run_pascal(r#"program T; type TInner=record V:Integer; end; type TOuter=record Tag:Integer; Inner:TInner; end; var o:TOuter; begin o.Tag:=1; o.Inner.V:=9; o.Inner.V:=0; WriteLn(o.Tag); WriteLn(o.Inner.V); end."#),
        &["1", "0"]
    );
}

#[test]
fn case_record_four_tags() {
    assert_eq!(
        run_pascal(r#"program T; type TOp=record case Code:Integer of 0:(Add:Integer); 1:(Sub:Integer); 2:(Mul:Integer); 3:(Div:Integer); end; var o:TOp; begin o.Code:=3; o.Div:=2; WriteLn(o.Div); end."#),
        &["2"]
    );
}

#[test]
fn record_method_uses_other_field() {
    assert_eq!(
        run_pascal(r#"program T; type TRect=record W,H:Integer; function Area:Integer; end; function TRect.Area:Integer; begin Result:=W*H; end; var r:TRect; begin r.W:=5; r.H:=6; WriteLn(r.Area); end."#),
        &["30"]
    );
}

#[test]
fn case_record_byte_fields_in_variant() {
    assert_eq!(
        run_pascal(r#"program T; type TBytes=record case Mode:Integer of 0:(A,B,C:Byte); 1:(WordVal:Integer); end; var b:TBytes; begin b.Mode:=0; b.A:=1; b.B:=2; b.C:=3; WriteLn(b.A+b.B+b.C); end."#),
        &["6"]
    );
}

#[test]
fn record_in_array_initialized_in_loop() {
    assert_eq!(
        run_pascal(r#"program T; type TCell=record V:Integer; end; var a:array[0..2] of TCell; i:Integer; begin for i:=0 to 2 do a[i].V:=i*i; WriteLn(a[2].V); end."#),
        &["4"]
    );
}

#[test]
fn nested_record_string_field_upper() {
    assert_eq!(
        run_pascal(r#"program T; type TInner=record Name:string; end; type TOuter=record Inner:TInner; end; var o:TOuter; begin o.Inner.Name:='vybe'; WriteLn(AnsiUpperCase(o.Inner.Name)); end."#),
        &["VYBE"]
    );
}

#[test]
fn case_record_dispatch_procedure() {
    assert_eq!(
        run_pascal(r#"program T; type TVal=record case K:Integer of 0:(I:Integer); 1:(S:string); end; procedure Show(const v:TVal); begin if v.K=0 then WriteLn(v.I) else WriteLn(v.S); end; var v:TVal; begin v.K:=0; v.I:=88; Show(v); end."#),
        &["88"]
    );
}

#[test]
fn record_char_array_field() {
    assert_eq!(
        run_pascal(r#"program T; type TBuf=record Ch:array[0..1] of Char; end; var b:TBuf; begin b.Ch[0]:='X'; b.Ch[1]:='Y'; WriteLn(b.Ch[0]); WriteLn(b.Ch[1]); end."#),
        &["X", "Y"]
    );
}

#[test]
fn case_record_bool_real_branch() {
    assert_eq!(
        run_pascal(r#"program T; type TVal=record case IsInt:Boolean of true:(I:Integer); false:(R:Double); end; var v:TVal; begin v.IsInt:=false; v.R:=3.5; WriteLn(Round(v.R)); end."#),
        &["4"]
    );
}

#[test]
fn record_nested_with_case_at_outer() {
    assert_eq!(
        run_pascal(r#"program T; type TInner=record V:Integer; end; type TOuter=record case Tag:Integer of 0:(I:Integer); 1:(Inner:TInner); end; var o:TOuter; begin o.Tag:=1; o.Inner.V:=13; WriteLn(o.Inner.V); end."#),
        &["13"]
    );
}

#[test]
fn record_two_methods_chain() {
    assert_eq!(
        run_pascal(r#"program T; type TAcc=record N:Integer; procedure Inc1; function Get:Integer; end; procedure TAcc.Inc1; begin N:=N+1; end; function TAcc.Get:Integer; begin Result:=N; end; var a:TAcc; begin a.N:=0; a.Inc1; a.Inc1; WriteLn(a.Get); end."#),
        &["2"]
    );
}

#[test]
fn case_record_tag_preserved_after_field_write() {
    assert_eq!(
        run_pascal(r#"program T; type TVal=record case K:Integer of 1:(A:Integer); 2:(B:Integer); end; var v:TVal; begin v.K:=1; v.A:=5; WriteLn(v.K); end."#),
        &["1"]
    );
}
