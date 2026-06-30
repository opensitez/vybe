/// Packed records and variant record edge cases — beyond test_packed_records.rs and test_variant_records.rs.
use super::helpers::run_pascal;

#[test]
fn packed_record_three_bytes_sequential() {
    assert_eq!(
        run_pascal(r#"program T; type TP=packed record A,B,C:Byte; end; var p:TP; begin p.A:=1; p.B:=2; p.C:=3; WriteLn(p.A+p.B+p.C); end."#),
        &["6"]
    );
}

#[test]
fn packed_record_assign_preserves_all_fields() {
    assert_eq!(
        run_pascal(r#"program T; type TP=packed record X,Y:Byte; end; var a,b:TP; begin a.X:=7; a.Y:=8; b:=a; WriteLn(b.X); WriteLn(b.Y); end."#),
        &["7", "8"]
    );
}

#[test]
fn packed_record_boolean_and_byte() {
    assert_eq!(
        run_pascal(r#"program T; type TF=packed record Flag:Boolean; Code:Byte; end; var f:TF; begin f.Flag:=true; f.Code:=42; if f.Flag then WriteLn(f.Code); end."#),
        &["42"]
    );
}

#[test]
fn packed_record_in_array() {
    assert_eq!(
        run_pascal(r#"program T; type TP=packed record V:Byte; end; var a:array[0..1] of TP; begin a[0].V:=5; a[1].V:=6; WriteLn(a[1].V); end."#),
        &["6"]
    );
}

#[test]
fn packed_record_nested_packed_inner() {
    assert_eq!(
        run_pascal(r#"program T; type TInner=packed record B:Byte; end; type TOuter=packed record Inner:TInner; end; var o:TOuter; begin o.Inner.B:=9; WriteLn(o.Inner.B); end."#),
        &["9"]
    );
}

#[test]
fn packed_record_case_on_byte_field() {
    assert_eq!(
        run_pascal(r#"program T; type TH=packed record Ver:Byte; end; var h:TH; begin h.Ver:=3; case h.Ver of 1:WriteLn('a'); 2:WriteLn('b'); 3:WriteLn('c'); else WriteLn('?'); end; end."#),
        &["c"]
    );
}

#[test]
fn packed_record_char_and_byte_pair() {
    assert_eq!(
        run_pascal(r#"program T; type TM=packed record Code:Byte; Ch:Char; end; var m:TM; begin m.Code:=65; m.Ch:='A'; WriteLn(m.Ch); end."#),
        &["A"]
    );
}

#[test]
fn packed_record_modify_one_field_keeps_other() {
    assert_eq!(
        run_pascal(r#"program T; type TP=packed record Lo,Hi:Byte; end; var p:TP; begin p.Lo:=1; p.Hi:=2; p.Lo:=9; WriteLn(p.Hi); end."#),
        &["2"]
    );
}

#[test]
fn packed_record_passed_to_procedure_by_var() {
    assert_eq!(
        run_pascal(r#"program T; type TP=packed record V:Byte; end; procedure IncP(var p:TP); begin p.V:=p.V+1; end; var p:TP; begin p.V:=4; IncP(p); WriteLn(p.V); end."#),
        &["5"]
    );
}

#[test]
fn packed_record_inside_regular_record() {
    assert_eq!(
        run_pascal(r#"program T; type TInner=packed record B:Byte; end; type TOuter=record Tag:Integer; Inner:TInner; end; var o:TOuter; begin o.Tag:=1; o.Inner.B:=6; WriteLn(o.Inner.B); end."#),
        &["6"]
    );
}

#[test]
fn variant_record_tag_switch_changes_active_field() {
    assert_eq!(
        run_pascal(r#"program T; type TVal=record case K:Integer of 0:(I:Integer); 1:(S:string); end; var v:TVal; begin v.K:=0; v.I:=10; v.K:=1; v.S:='x'; WriteLn(v.S); end."#),
        &["x"]
    );
}

#[test]
fn variant_record_overlapping_bytes_and_word() {
    assert_eq!(
        run_pascal(r#"program T; type TW=record case T:Integer of 0:(Lo,Hi:Byte); 1:(W:Integer); end; var w:TW; begin w.T:=1; w.W:=258; WriteLn(w.W); end."#),
        &["258"]
    );
}

#[test]
fn variant_record_enum_two_arms() {
    assert_eq!(
        run_pascal(r#"program T; type TK=(A,B); type TV=record case Kind:TK of A:(N:Integer); B:(T:string); end; var v:TV; begin v.Kind:=B; v.T:='ok'; WriteLn(v.T); end."#),
        &["ok"]
    );
}

#[test]
fn variant_record_fixed_header_plus_variant() {
    assert_eq!(
        run_pascal(r#"program T; type TMsg=record Id:Integer; case Kind:Integer of 1:(Text:string); 2:(Code:Integer); end; var m:TMsg; begin m.Id:=1; m.Kind:=1; m.Text:='hi'; WriteLn(m.Text); end."#),
        &["hi"]
    );
}

#[test]
fn variant_record_three_arms_dispatch() {
    assert_eq!(
        run_pascal(r#"program T; type TOp=record case C:Integer of 0:(A:Integer); 1:(B:Integer); 2:(Cval:Integer); end; procedure Show(const o:TOp); begin case o.C of 0:WriteLn(o.A); 1:WriteLn(o.B); 2:WriteLn(o.Cval); end; end; var x:TOp; begin x.C:=2; x.Cval:=99; Show(x); end."#),
        &["99"]
    );
}

#[test]
fn variant_record_boolean_arms_real_and_int() {
    assert_eq!(
        run_pascal(r#"program T; type TN=record case IsReal:Boolean of false:(I:Integer); true:(R:Double); end; var n:TN; begin n.IsReal:=true; n.R:=2.5; WriteLn(Round(n.R)); end."#),
        &["3"]
    );
}

#[test]
fn variant_record_nested_fixed_in_arm() {
    assert_eq!(
        run_pascal(r#"program T; type TInner=record V:Integer; end; type TWrap=record case K:Integer of 0:(I:Integer); 1:(Inner:TInner); end; var w:TWrap; begin w.K:=1; w.Inner.V:=44; WriteLn(w.Inner.V); end."#),
        &["44"]
    );
}

#[test]
fn packed_variant_combined_tag_byte() {
    assert_eq!(
        run_pascal(r#"program T; type TH=packed record case Tag:Byte of 0:(A:Byte); 1:(B:Byte); end; var h:TH; begin h.Tag:=1; h.B:=7; WriteLn(h.B); end."#),
        &["7"]
    );
}

#[test]
fn variant_record_char_tag_arms() {
    assert_eq!(
        run_pascal(r#"program T; type TT=record case K:Char of 'N':(I:Integer); 'S':(T:string); end; var v:TT; begin v.K:='N'; v.I:=12; WriteLn(v.I); end."#),
        &["12"]
    );
}

#[test]
fn variant_record_multiple_labels_one_arm() {
    assert_eq!(
        run_pascal(r#"program T; type TV=record case K:Integer of 0,1:(X:Integer); 2:(Y:Integer); end; var v:TV; begin v.K:=1; v.X:=55; WriteLn(v.X); end."#),
        &["55"]
    );
}

#[test]
fn packed_record_loop_accumulate() {
    assert_eq!(
        run_pascal(r#"program T; type TP=packed record V:Byte; end; var a:array[0..2] of TP; i,s:Integer; begin for i:=0 to 2 do begin a[i].V:=i+1; s:=0; end; for i:=0 to 2 do s:=s+a[i].V; WriteLn(s); end."#),
        &["6"]
    );
}

#[test]
fn variant_record_if_on_tag_before_read() {
    assert_eq!(
        run_pascal(r#"program T; type TVal=record case K:Integer of 0:(I:Integer); 1:(S:string); end; var v:TVal; begin v.K:=0; v.I:=8; if v.K=0 then WriteLn(v.I) else WriteLn(0); end."#),
        &["8"]
    );
}

#[test]
fn packed_record_compare_two_instances() {
    assert_eq!(
        run_pascal(r#"program T; type TP=packed record V:Byte; end; var a,b:TP; begin a.V:=3; b.V:=3; WriteLn(a.V=b.V); end."#),
        &["true"]
    );
}

#[test]
fn variant_record_string_arm_length() {
    assert_eq!(
        run_pascal(r#"program T; type TVal=record case K:Integer of 0:(I:Integer); 1:(S:string); end; var v:TVal; begin v.K:=1; v.S:='abcd'; WriteLn(Length(v.S)); end."#),
        &["4"]
    );
}

#[test]
fn packed_record_shortint_field() {
    assert_eq!(
        run_pascal(r#"program T; type TP=packed record V:ShortInt; end; var p:TP; begin p.V:=-5; WriteLn(p.V); end."#),
        &["-5"]
    );
}

#[test]
fn variant_record_write_int_then_read_after_tag_unchanged() {
    assert_eq!(
        run_pascal(r#"program T; type TVal=record case K:Integer of 1:(A:Integer); 2:(B:Integer); end; var v:TVal; begin v.K:=1; v.A:=77; WriteLn(v.K); end."#),
        &["1"]
    );
}

#[test]
fn packed_record_used_in_function_result() {
    assert_eq!(
        run_pascal(r#"program T; type TP=packed record V:Byte; end; function Make(v:Byte):TP; begin Result.V:=v; end; var p:TP; begin p:=Make(11); WriteLn(p.V); end."#),
        &["11"]
    );
}

#[test]
fn variant_record_case_statement_on_tag() {
    assert_eq!(
        run_pascal(r#"program T; type TVal=record case K:Integer of 0:(I:Integer); 1:(S:string); end; var v:TVal; begin v.K:=0; v.I:=6; case v.K of 0:WriteLn(v.I); 1:WriteLn(0); end; end."#),
        &["6"]
    );
}

#[test]
fn packed_record_four_bytes_sum() {
    assert_eq!(
        run_pascal(r#"program T; type TP=packed record A,B,C,D:Byte; end; var p:TP; begin p.A:=1; p.B:=2; p.C:=3; p.D:=4; WriteLn(p.A+p.B+p.C+p.D); end."#),
        &["10"]
    );
}

#[test]
fn variant_record_real_arm_fraction() {
    assert_eq!(
        run_pascal(r#"program T; type TVal=record case K:Integer of 0:(I:Integer); 1:(R:Double); end; var v:TVal; begin v.K:=1; v.R:=1.25; WriteLn(Frac(v.R)=0.25); end."#),
        &["true"]
    );
}

#[test]
fn packed_record_reset_field_zero() {
    assert_eq!(
        run_pascal(r#"program T; type TP=packed record V:Byte; end; var p:TP; begin p.V:=9; p.V:=0; WriteLn(p.V); end."#),
        &["0"]
    );
}

#[test]
fn variant_record_two_fixed_before_case() {
    assert_eq!(
        run_pascal(r#"program T; type TMsg=record Ver,Id:Integer; case Kind:Integer of 1:(Text:string); 2:(Code:Integer); end; var m:TMsg; begin m.Ver:=1; m.Id:=9; m.Kind:=2; m.Code:=500; WriteLn(m.Code); end."#),
        &["500"]
    );
}

#[test]
fn packed_record_array_copy_element() {
    assert_eq!(
        run_pascal(r#"program T; type TP=packed record V:Byte; end; var a,b:array[0..0] of TP; begin a[0].V:=3; b[0]:=a[0]; WriteLn(b[0].V); end."#),
        &["3"]
    );
}

#[test]
fn variant_record_procedure_var_param() {
    assert_eq!(
        run_pascal(r#"program T; type TVal=record case K:Integer of 0:(I:Integer); 1:(S:string); end; procedure SetInt(var v:TVal; n:Integer); begin v.K:=0; v.I:=n; end; var v:TVal; begin SetInt(v,5); WriteLn(v.I); end."#),
        &["5"]
    );
}

#[test]
fn packed_record_word_byte_mixed() {
    assert_eq!(
        run_pascal(r#"program T; type TP=packed record Lo:Byte; Hi:Byte; end; var p:TP; begin p.Lo:=1; p.Hi:=2; WriteLn(p.Lo*10+p.Hi); end."#),
        &["12"]
    );
}

#[test]
fn variant_record_tag_2_byte_fields_sum() {
    assert_eq!(
        run_pascal(r#"program T; type TPair=record case K:Integer of 0:(A,B:Byte); 1:(V:Integer); end; var p:TPair; begin p.K:=0; p.A:=2; p.B:=3; WriteLn(p.A+p.B); end."#),
        &["5"]
    );
}

#[test]
fn packed_record_in_case_variant_outer() {
    assert_eq!(
        run_pascal(r#"program T; type TInner=packed record B:Byte; end; type TOuter=record case Tag:Integer of 0:(I:Integer); 1:(Inner:TInner); end; var o:TOuter; begin o.Tag:=1; o.Inner.B:=8; WriteLn(o.Inner.B); end."#),
        &["8"]
    );
}

#[test]
fn variant_record_bool_false_int_arm() {
    assert_eq!(
        run_pascal(r#"program T; type TVal=record case B:Boolean of false:(N:Integer); true:(S:string); end; var v:TVal; begin v.B:=false; v.N:=123; WriteLn(v.N); end."#),
        &["123"]
    );
}

#[test]
fn packed_record_logic_not_flag() {
    assert_eq!(
        run_pascal(r#"program T; type TF=packed record On:Boolean; end; var f:TF; begin f.On:=true; f.On:=not f.On; WriteLn(f.On); end."#),
        &["false"]
    );
}

#[test]
fn variant_record_four_way_tag_dispatch() {
    assert_eq!(
        run_pascal(r#"program T; type TOp=record case C:Integer of 0:(A:Integer); 1:(B:Integer); 2:(C2:Integer); 3:(D:Integer); end; var o:TOp; begin o.C:=3; o.D:=21; WriteLn(o.D); end."#),
        &["21"]
    );
}
