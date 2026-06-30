/// Subrange types: bounds, assignment, and ordinal use.
use super::helpers::run_pascal;

#[test]
fn subrange_percent_75() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPercent=0..100; var p:TPercent; begin p:=75; WriteLn(p); end."#
        ),
        &["75"]
    );
}

#[test]
fn subrange_percent_zero() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPercent=0..100; var p:TPercent; begin p:=0; WriteLn(p); end."#
        ),
        &["0"]
    );
}

#[test]
fn subrange_percent_max() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPercent=0..100; var p:TPercent; begin p:=100; WriteLn(p); end."#
        ),
        &["100"]
    );
}

#[test]
fn subrange_small_negative() {
    assert_eq!(
        run_pascal(
            r#"program T; type TSmall=-5..5; var x:TSmall; begin x:=-3; WriteLn(x); end."#
        ),
        &["-3"]
    );
}

#[test]
fn subrange_small_positive() {
    assert_eq!(
        run_pascal(
            r#"program T; type TSmall=-5..5; var x:TSmall; begin x:=4; WriteLn(x); end."#
        ),
        &["4"]
    );
}

#[test]
fn subrange_month_december() {
    assert_eq!(
        run_pascal(
            r#"program T; type TMonth=1..12; var m:TMonth; begin m:=12; WriteLn(m); end."#
        ),
        &["12"]
    );
}

#[test]
fn subrange_month_january() {
    assert_eq!(
        run_pascal(
            r#"program T; type TMonth=1..12; var m:TMonth; begin m:=1; WriteLn(m); end."#
        ),
        &["1"]
    );
}

#[test]
fn subrange_day_hour() {
    assert_eq!(
        run_pascal(
            r#"program T; type THour=0..23; var h:THour; begin h:=23; WriteLn(h); end."#
        ),
        &["23"]
    );
}

#[test]
fn subrange_day_minute() {
    assert_eq!(
        run_pascal(
            r#"program T; type TMinute=0..59; var m:TMinute; begin m:=59; WriteLn(m); end."#
        ),
        &["59"]
    );
}

#[test]
fn subrange_char_letter_m() {
    assert_eq!(
        run_pascal(
            r#"program T; type TLetter='a'..'z'; var c:TLetter; begin c:='m'; WriteLn(c); end."#
        ),
        &["m"]
    );
}

#[test]
fn subrange_char_letter_a() {
    assert_eq!(
        run_pascal(
            r#"program T; type TLetter='a'..'z'; var c:TLetter; begin c:='a'; WriteLn(c); end."#
        ),
        &["a"]
    );
}

#[test]
fn subrange_char_digit_5() {
    assert_eq!(
        run_pascal(
            r#"program T; type TDigit='0'..'9'; var d:TDigit; begin d:='5'; WriteLn(d); end."#
        ),
        &["5"]
    );
}

#[test]
fn subrange_char_digit_0() {
    assert_eq!(
        run_pascal(
            r#"program T; type TDigit='0'..'9'; var d:TDigit; begin d:='0'; WriteLn(d); end."#
        ),
        &["0"]
    );
}

#[test]
fn subrange_assign_from_var() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=1..10; var a,b:TR; begin a:=3; b:=a; WriteLn(b); end."#
        ),
        &["3"]
    );
}

#[test]
fn subrange_inc_in_loop() {
    assert_eq!(
        run_pascal(
            r#"program T; type TIdx=1..4; var i:TIdx; begin for i:=1 to 4 do WriteLn(i); end."#
        ),
        &["1", "2", "3", "4"]
    );
}

#[test]
fn subrange_succ_builtin() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=10..20; var x:TR; begin x:=15; WriteLn(Succ(x)); end."#
        ),
        &["16"]
    );
}

#[test]
fn subrange_pred_builtin() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=10..20; var x:TR; begin x:=15; WriteLn(Pred(x)); end."#
        ),
        &["14"]
    );
}

#[test]
fn subrange_ord_char() {
    assert_eq!(
        run_pascal(
            r#"program T; type TLetter='A'..'Z'; var c:TLetter; begin c:='C'; WriteLn(Ord(c)); end."#
        ),
        &["67"]
    );
}

#[test]
fn subrange_set_membership() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=1..3; var s:set of TD; x:TD; begin s:=[1,3]; x:=2; if x in s then WriteLn('in') else WriteLn('out'); end."#
        ),
        &["out"]
    );
}

#[test]
fn subrange_set_include() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=1..3; var s:set of TD; begin s:=[2]; if 2 in s then WriteLn('yes'); end."#
        ),
        &["yes"]
    );
}

#[test]
fn subrange_add_to_integer() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=1..5; var x:TR; n:Integer; begin x:=3; n:=x+2; WriteLn(n); end."#
        ),
        &["5"]
    );
}

#[test]
fn subrange_compare_equal() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=0..9; var a,b:TR; begin a:=7; b:=7; if a=b then WriteLn('eq'); end."#
        ),
        &["eq"]
    );
}

#[test]
fn subrange_compare_less() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=0..9; var a,b:TR; begin a:=2; b:=8; if a<b then WriteLn('lt'); end."#
        ),
        &["lt"]
    );
}

#[test]
fn subrange_byte_range() {
    assert_eq!(
        run_pascal(
            r#"program T; type TByteRange=0..255; var b:TByteRange; begin b:=255; WriteLn(b); end."#
        ),
        &["255"]
    );
}

#[test]
fn subrange_narrow_1_3() {
    assert_eq!(
        run_pascal(
            r#"program T; type TN=1..3; var n:TN; begin n:=2; WriteLn(n*2); end."#
        ),
        &["4"]
    );
}

#[test]
fn subrange_negative_bound() {
    assert_eq!(
        run_pascal(
            r#"program T; type TNeg=-10..-1; var n:TNeg; begin n:=-1; WriteLn(n); end."#
        ),
        &["-1"]
    );
}

#[test]
fn subrange_char_upper_z() {
    assert_eq!(
        run_pascal(
            r#"program T; type TZ='A'..'Z'; var c:TZ; begin c:='Z'; WriteLn(c); end."#
        ),
        &["Z"]
    );
}

#[test]
fn subrange_for_downto() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=3..1; var i:Integer; begin for i:=3 downto 1 do WriteLn(i); end."#
        ),
        &["3", "2", "1"]
    );
}

#[test]
fn subrange_record_field() {
    assert_eq!(
        run_pascal(
            r#"program T; type TDay=1..31; type TDate=record D:TDay; end; var dt:TDate; begin dt.D:=15; WriteLn(dt.D); end."#
        ),
        &["15"]
    );
}

#[test]
fn subrange_array_index() {
    assert_eq!(
        run_pascal(
            r#"program T; type TI=1..3; type TA=array[TI] of Integer; var a:TA; i:TI; begin for i:=1 to 3 do a[i]:=i*10; WriteLn(a[2]); end."#
        ),
        &["20"]
    );
}

#[test]
fn subrange_case_label() {
    assert_eq!(
        run_pascal(
            r#"program T; type TM=1..3; var m:TM; begin m:=2; case m of 1:WriteLn('a'); 2:WriteLn('b'); 3:WriteLn('c'); end; end."#
        ),
        &["b"]
    );
}

#[test]
fn subrange_low_high_style() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=5..9; var r:TR; begin r:=5; WriteLn(r); r:=9; WriteLn(r); end."#
        ),
        &["5", "9"]
    );
}

#[test]
fn subrange_assign_literal_min() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=100..200; var x:TR; begin x:=100; WriteLn(x); end."#
        ),
        &["100"]
    );
}

#[test]
fn subrange_assign_literal_max() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=100..200; var x:TR; begin x:=200; WriteLn(x); end."#
        ),
        &["200"]
    );
}

#[test]
fn subrange_char_assign_from_var() {
    assert_eq!(
        run_pascal(
            r#"program T; type TL='a'..'c'; var a,b:TL; begin a:='b'; b:=a; WriteLn(b); end."#
        ),
        &["b"]
    );
}

#[test]
fn subrange_in_procedure_param() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=1..5; procedure P(v:TR); begin WriteLn(v); end; begin P(4); end."#
        ),
        &["4"]
    );
}

#[test]
fn subrange_sum_in_expr() {
    assert_eq!(
        run_pascal(
            r#"program T; type TA=1..3; type TB=4..6; var a:TA; b:Integer; begin a:=2; b:=a+4; WriteLn(b); end."#
        ),
        &["6"]
    );
}

#[test]
fn subrange_mod_result() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=0..9; var x:TR; begin x:=7; WriteLn(x mod 4); end."#
        ),
        &["3"]
    );
}

#[test]
fn subrange_multiply() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=2..5; var x:TR; begin x:=3; WriteLn(x*3); end."#
        ),
        &["9"]
    );
}

#[test]
fn subrange_boolean_from_compare() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=1..10; var x:TR; begin x:=6; if x>5 then WriteLn('gt') else WriteLn('le'); end."#
        ),
        &["gt"]
    );
}

#[test]
fn subrange_nested_type() {
    assert_eq!(
        run_pascal(
            r#"program T; type TInner=1..2; type TOuter=record V:TInner; end; var o:TOuter; begin o.V:=2; WriteLn(o.V); end."#
        ),
        &["2"]
    );
}

#[test]
fn subrange_zero_based() {
    assert_eq!(
        run_pascal(
            r#"program T; type TIdx=0..4; var i:TIdx; s:Integer; begin s:=0; for i:=0 to 4 do s:=s+i; WriteLn(s); end."#
        ),
        &["10"]
    );
}

