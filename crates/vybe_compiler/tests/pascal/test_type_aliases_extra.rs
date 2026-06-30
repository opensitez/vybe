/// Type aliases including nested alias chains.
use super::helpers::run_pascal;

#[test]
fn alias_integer_count() {
    assert_eq!(
        run_pascal(
            r#"program T; type TCount=Integer; var v:TCount; begin v:=7; WriteLn(v); end."#
        ),
        &["7"]
    );
}

#[test]
fn alias_string_name() {
    assert_eq!(
        run_pascal(
            r#"program T; type TName=string; var v:TName; begin v:='bob'; WriteLn(v); end."#
        ),
        &["bob"]
    );
}

#[test]
fn alias_boolean_flag() {
    assert_eq!(
        run_pascal(
            r#"program T; type TFlag=Boolean; var v:TFlag; begin v:=true; WriteLn(v); end."#
        ),
        &["true"]
    );
}

#[test]
fn alias_char_code() {
    assert_eq!(
        run_pascal(
            r#"program T; type TCode=Char; var v:TCode; begin v:='X'; WriteLn(v); end."#
        ),
        &["X"]
    );
}

#[test]
fn alias_byte_small() {
    assert_eq!(
        run_pascal(
            r#"program T; type TByteAlias=Byte; var v:TByteAlias; begin v:=200; WriteLn(v); end."#
        ),
        &["200"]
    );
}

#[test]
fn alias_real_amount() {
    assert_eq!(
        run_pascal(
            r#"program T; type TAmount=Real; var v:TAmount; begin v:=3.5; WriteLn(v); end."#
        ),
        &["3.5"]
    );
}

#[test]
fn alias_nested_id_in_record() {
    assert_eq!(
        run_pascal(
            r#"program T; type TId=Integer; type TRec=record Id:TId; end; var r:TRec; begin r.Id:=42; WriteLn(r.Id); end."#
        ),
        &["42"]
    );
}

#[test]
fn alias_double_pointer() {
    assert_eq!(
        run_pascal(
            r#"program T; type TInt=Integer; type PInt=^TInt; var n:TInt; p:PInt; begin n:=99; p:=@n; WriteLn(p^); end."#
        ),
        &["99"]
    );
}

#[test]
fn alias_array_of_int() {
    assert_eq!(
        run_pascal(
            r#"program T; type TInt=Integer; type TArr=array[0..2] of TInt; var a:TArr; begin a[1]:=5; WriteLn(a[1]); end."#
        ),
        &["5"]
    );
}

#[test]
fn alias_string_array() {
    assert_eq!(
        run_pascal(
            r#"program T; type TLine=string; type TLines=array[1..2] of TLine; var L:TLines; begin L[1]:='a'; L[2]:='b'; WriteLn(L[1]); WriteLn(L[2]); end."#
        ),
        &["a", "b"]
    );
}

#[test]
fn alias_chain_a_b_c() {
    assert_eq!(
        run_pascal(
            r#"program T; type A=Integer; type B=A; type C=B; var x:C; begin x:=11; WriteLn(x); end."#
        ),
        &["11"]
    );
}

#[test]
fn alias_chain_string() {
    assert_eq!(
        run_pascal(
            r#"program T; type S1=string; type S2=S1; type S3=S2; var t:S3; begin t:='nest'; WriteLn(t); end."#
        ),
        &["nest"]
    );
}

#[test]
fn alias_record_field_alias() {
    assert_eq!(
        run_pascal(
            r#"program T; type TScore=Integer; type TPlayer=record Score:TScore; end; var p:TPlayer; begin p.Score:=100; WriteLn(p.Score); end."#
        ),
        &["100"]
    );
}

#[test]
fn alias_enum_underlying() {
    assert_eq!(
        run_pascal(
            r#"program T; type TKind=(A,B,C); type TRef=TKind; var k:TRef; begin k:=B; WriteLn(Ord(k)); end."#
        ),
        &["1"]
    );
}

#[test]
fn alias_pointer_to_alias() {
    assert_eq!(
        run_pascal(
            r#"program T; type TNum=Integer; type PNum=^TNum; var n:TNum; p:PNum; begin n:=7; p:=@n; WriteLn(p^); end."#
        ),
        &["7"]
    );
}

#[test]
fn alias_set_of_alias() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=1..3; type TSet=set of TD; var s:TSet; begin s:=[1,3]; if 2 in s then WriteLn('in') else WriteLn('out'); end."#
        ),
        &["out"]
    );
}

#[test]
fn alias_subrange_under_int() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=1..5; type TIdx=TR; var i:TIdx; begin i:=4; WriteLn(i); end."#
        ),
        &["4"]
    );
}

#[test]
fn alias_dynamic_array() {
    assert_eq!(
        run_pascal(
            r#"program T; type TE=Integer; var a:array of TE; begin SetLength(a,2); a[0]:=1; a[1]:=2; WriteLn(a[1]); end."#
        ),
        &["2"]
    );
}

#[test]
fn alias_proc_type() {
    assert_eq!(
        run_pascal(
            r#"program T; type TInt=Integer; type TFn=function(x:TInt):TInt; function DoubleIt(n:TInt):TInt; begin Result:=n*2; end; var f:TFn; begin f:=@DoubleIt; WriteLn(f(6)); end."#
        ),
        &["12"]
    );
}

#[test]
fn alias_nested_record_alias() {
    assert_eq!(
        run_pascal(
            r#"program T; type TInner=record V:Integer; end; type TOuter=record Inner:TInner; end; var o:TOuter; begin o.Inner.V:=3; WriteLn(o.Inner.V); end."#
        ),
        &["3"]
    );
}

#[test]
fn alias_word_type() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPort=Word; var p:TPort; begin p:=8080; WriteLn(p); end."#
        ),
        &["8080"]
    );
}

#[test]
fn alias_longint() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBig=Int64; var n:TBig; begin n:=1000000; WriteLn(n); end."#
        ),
        &["1000000"]
    );
}

#[test]
fn alias_currency_style() {
    assert_eq!(
        run_pascal(
            r#"program T; type TCents=Integer; type TDollars=Real; var c:TCents; d:TDollars; begin c:=250; d:=c/100; WriteLn(d); end."#
        ),
        &["2.5"]
    );
}

#[test]
fn alias_char_set() {
    assert_eq!(
        run_pascal(
            r#"program T; type TLetter='a'..'z'; var c:TLetter; begin c:='k'; WriteLn(c); end."#
        ),
        &["k"]
    );
}

#[test]
fn alias_two_d_array() {
    assert_eq!(
        run_pascal(
            r#"program T; type TCell=Integer; type TGrid=array[0..1,0..1] of TCell; var g:TGrid; begin g[0,1]:=9; WriteLn(g[0,1]); end."#
        ),
        &["9"]
    );
}

#[test]
fn alias_const_via_type() {
    assert_eq!(
        run_pascal(
            r#"program T; type TCode=Integer; const C:TCode=42; var x:TCode; begin x:=C; WriteLn(x); end."#
        ),
        &["42"]
    );
}

#[test]
fn alias_string_pointer() {
    assert_eq!(
        run_pascal(
            r#"program T; type TStr=string; type PStr=^TStr; var s:TStr; p:PStr; begin s:='hi'; p:=@s; WriteLn(p^); end."#
        ),
        &["hi"]
    );
}

#[test]
fn alias_nested_twice_record() {
    assert_eq!(
        run_pascal(
            r#"program T; type TA=Integer; type TB=record A:TA; end; type TC=record B:TB; end; var c:TC; begin c.B.A:=8; WriteLn(c.B.A); end."#
        ),
        &["8"]
    );
}

#[test]
fn alias_enum_array() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=(X,Y,Z); type TArr=array[T] of string; var a:TArr; begin a[Y]:='mid'; WriteLn(a[Y]); end."#
        ),
        &["mid"]
    );
}

#[test]
fn alias_ref_to_subrange() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=10..20; type TPtr=^TR; var v:TR; p:TPtr; begin v:=15; p:=@v; WriteLn(p^); end."#
        ),
        &["15"]
    );
}

#[test]
fn alias_multiline_type_block() {
    assert_eq!(
        run_pascal(
            r#"program T; type TInt=Integer; type TPair=record A,B:TInt; end; var p:TPair; begin p.A:=2; p.B:=3; WriteLn(p.A+p.B); end."#
        ),
        &["5"]
    );
}

#[test]
fn alias_string_sub() {
    assert_eq!(
        run_pascal(
            r#"program T; type TTag=string; type TName=TTag; var n:TName; begin n:='tag'; WriteLn(Length(n)); end."#
        ),
        &["3"]
    );
}

#[test]
fn alias_bool_not() {
    assert_eq!(
        run_pascal(
            r#"program T; type TSwitch=Boolean; var s:TSwitch; begin s:=false; WriteLn(not s); end."#
        ),
        &["true"]
    );
}

#[test]
fn alias_real_trunc() {
    assert_eq!(
        run_pascal(
            r#"program T; type TVal=Real; var v:TVal; begin v:=9.7; WriteLn(Trunc(v)); end."#
        ),
        &["9"]
    );
}

#[test]
fn alias_char_ord() {
    assert_eq!(
        run_pascal(
            r#"program T; type TCh=Char; var c:TCh; begin c:='B'; WriteLn(Ord(c)); end."#
        ),
        &["66"]
    );
}

#[test]
fn alias_integer_hex() {
    assert_eq!(
        run_pascal(
            r#"program T; type THex=Integer; const H:THex=$10; begin WriteLn(H); end."#
        ),
        &["16"]
    );
}

#[test]
fn alias_pointer_chain() {
    assert_eq!(
        run_pascal(
            r#"program T; type TI=Integer; type PI=^TI; type PPI=^PI; var n:TI; p:PI; pp:PPI; begin n:=5; p:=@n; pp:=@p; WriteLn(pp^^); end."#
        ),
        &["5"]
    );
}

#[test]
fn alias_record_array_field() {
    assert_eq!(
        run_pascal(
            r#"program T; type TItem=Integer; type TBag=record Items:array[0..1] of TItem; end; var b:TBag; begin b.Items[1]:=4; WriteLn(b.Items[1]); end."#
        ),
        &["4"]
    );
}

#[test]
fn alias_nested_string_alias() {
    assert_eq!(
        run_pascal(
            r#"program T; type S=string; type SS=S; type SSS=SS; var t:SSS; begin t:='deep'; WriteLn(t); end."#
        ),
        &["deep"]
    );
}

#[test]
fn alias_func_result_type() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=Integer; function F:TR; begin Result:=21; end; begin WriteLn(F); end."#
        ),
        &["21"]
    );
}

#[test]
fn alias_set_enum() {
    assert_eq!(
        run_pascal(
            r#"program T; type TC=(R,G,B); type TS=set of TC; var s:TS; begin s:=[R,B]; if G in s then WriteLn('g') else WriteLn('no'); end."#
        ),
        &["no"]
    );
}

#[test]
fn alias_variant_style_integer() {
    assert_eq!(
        run_pascal(
            r#"program T; type TMaybe=Integer; var m:TMaybe; begin m:=0; WriteLn(m); end."#
        ),
        &["0"]
    );
}

