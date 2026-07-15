/// Static (fixed) array declarations, bounds, indexing, and passing.
use super::helpers::run_pascal;

#[test]
fn static_array_low_high_bounds() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array[1..4] of Integer; begin WriteLn(Low(a)); WriteLn(High(a)); end."#
        ),
        &["1", "4"]
    );
}

#[test]
fn static_array_zero_based_indexing() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array[0..2] of Integer; begin a[0]:=10; a[2]:=30; WriteLn(a[0]+a[2]); end."#
        ),
        &["40"]
    );
}

#[test]
fn static_array_negative_lower_bound() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array[-1..1] of Integer; begin a[-1]:=5; a[1]:=7; WriteLn(a[-1]+a[1]); end."#
        ),
        &["12"]
    );
}

#[test]
fn static_array_literal_initializer() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array[0..2] of Integer = (1,2,3); begin WriteLn(a[1]); end."#
        ),
        &["2"]
    );
}

#[test]
fn static_array_length_via_high_low() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array[2..6] of Integer; begin WriteLn(High(a)-Low(a)+1); end."#
        ),
        &["5"]
    );
}

#[test]
fn static_array_char_elements_join() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array[1..3] of Char; s: string; begin a[1]:='a'; a[2]:='b'; a[3]:='c'; s:=a[1]+a[2]+a[3]; WriteLn(s); end."#
        ),
        &["abc"]
    );
}

#[test]
fn static_array_boolean_flags_all_true() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array[1..3] of Boolean; i: Integer; ok: Boolean; begin for i:=1 to 3 do a[i]:=true; ok:=a[1] and a[2] and a[3]; WriteLn(ok); end."#
        ),
        &["true"]
    );
}

#[test]
fn static_array_real_sum_loop() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array[1..3] of Double; i: Integer; s: Double; begin a[1]:=1.0; a[2]:=2.0; a[3]:=3.0; s:=0; for i:=1 to 3 do s:=s+a[i]; WriteLn(Round(s)); end."#
        ),
        &["6"]
    );
}

#[test]
fn static_array_swap_elements() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array[1..2] of Integer; t: Integer; begin a[1]:=1; a[2]:=9; t:=a[1]; a[1]:=a[2]; a[2]:=t; WriteLn(a[1]); WriteLn(a[2]); end."#
        ),
        &["9", "1"]
    );
}

#[test]
fn static_array_copy_loop() {
    assert_eq!(
        run_pascal(
            r#"program T; var src,dst: array[0..1] of Integer; i: Integer; begin src[0]:=4; src[1]:=5; for i:=0 to 1 do dst[i]:=src[i]; WriteLn(dst[0]+dst[1]); end."#
        ),
        &["9"]
    );
}

#[test]
fn static_array_open_parameter_sum() {
    assert_eq!(
        run_pascal(
            r#"program T; function Sum3(a: array[0..2] of Integer): Integer; begin Result:=a[0]+a[1]+a[2]; end; var v: array[0..2] of Integer; begin v[0]:=1; v[1]:=2; v[2]:=3; WriteLn(Sum3(v)); end."#
        ),
        &["6"]
    );
}

#[test]
fn static_array_nested_type_field() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; end; var a: array[1..2] of TR; begin a[1].V:=3; a[2].V:=4; WriteLn(a[1].V+a[2].V); end."#
        ),
        &["7"]
    );
}

#[test]
fn static_array_enum_indexed() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=(A,B,C); var arr: array[TD] of Integer; begin arr[A]:=1; arr[C]:=3; WriteLn(arr[A]+arr[C]); end."#
        ),
        &["4"]
    );
}

#[test]
fn static_array_string_elements() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array[1..2] of string; begin a[1]:='hi'; a[2]:='!'; WriteLn(a[1]+a[2]); end."#
        ),
        &["hi!"]
    );
}

#[test]
fn static_array_find_max_value() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array[1..4] of Integer; i,m: Integer; begin a[1]:=3; a[2]:=9; a[3]:=1; a[4]:=7; m:=a[1]; for i:=2 to 4 do if a[i]>m then m:=a[i]; WriteLn(m); end."#
        ),
        &["9"]
    );
}

#[test]
fn static_array_reverse_in_place() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array[1..3] of Integer; i,t: Integer; begin a[1]:=1; a[2]:=2; a[3]:=3; for i:=1 to 1 do begin t:=a[i]; a[i]:=a[4-i]; a[4-i]:=t; end; WriteLn(a[2]); end."#
        ),
        &["2"]
    );
}

#[test]
fn static_array_count_matching() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array[1..5] of Integer; i,c: Integer; begin a[1]:=2; a[2]:=2; a[3]:=5; a[4]:=2; a[5]:=1; c:=0; for i:=1 to 5 do if a[i]=2 then c:=c+1; WriteLn(c); end."#
        ),
        &["3"]
    );
}

#[test]
fn static_array_multiply_accumulate() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array[1..3] of Integer; i,p: Integer; begin a[1]:=2; a[2]:=3; a[3]:=4; p:=1; for i:=1 to 3 do p:=p*a[i]; WriteLn(p); end."#
        ),
        &["24"]
    );
}

#[test]
fn static_array_shift_left_one() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array[0..2] of Integer; i: Integer; begin a[0]:=1; a[1]:=2; a[2]:=3; for i:=0 to 1 do a[i]:=a[i+1]; a[2]:=0; WriteLn(a[0]); WriteLn(a[2]); end."#
        ),
        &["2", "0"]
    );
}

#[test]
fn static_array_constant_index_expression() {
    assert_eq!(
        run_pascal(
            r#"program T; const k=1; var a: array[1..3] of Integer; begin a[k+1]:=8; WriteLn(a[2]); end."#
        ),
        &["8"]
    );
}

#[test]
fn static_array_of_byte_ord_sum() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array[0..1] of Byte; begin a[0]:=10; a[1]:=20; WriteLn(a[0]+a[1]); end."#
        ),
        &["30"]
    );
}

#[test]
fn static_array_partial_fill_then_read() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array[1..4] of Integer; i: Integer; begin for i:=1 to 4 do a[i]:=i*i; WriteLn(a[3]); end."#
        ),
        &["9"]
    );
}

#[test]
fn static_array_compare_first_last() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array[1..3] of Integer; begin a[1]:=5; a[3]:=5; if a[1]=a[3] then WriteLn('eq'); end."#
        ),
        &["eq"]
    );
}

#[test]
fn static_array_nested_in_record() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=record Items: array[0..1] of Integer; end; var b: TB; begin b.Items[0]:=6; b.Items[1]:=7; WriteLn(b.Items[0]+b.Items[1]); end."#
        ),
        &["13"]
    );
}

#[test]
fn static_array_class_field() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBox=class public Data: array[1..2] of Integer; end; var b: TBox; begin b:=TBox.Create; b.Data[1]:=11; b.Data[2]:=22; WriteLn(b.Data[1]+b.Data[2]); b.Free; end."#
        ),
        &["33"]
    );
}

#[test]
fn static_array_procedure_var_param() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure DoubleAll(var a: array[0..1] of Integer); begin a[0]:=a[0]*2; a[1]:=a[1]*2; end; var v: array[0..1] of Integer; begin v[0]:=3; v[1]:=4; DoubleAll(v); WriteLn(v[0]+v[1]); end."#
        ),
        &["14"]
    );
}

#[test]
fn static_array_function_returns_element() {
    assert_eq!(
        run_pascal(
            r#"program T; function Pick(a: array[1..3] of Integer; idx: Integer): Integer; begin Result:=a[idx]; end; var v: array[1..3] of Integer; begin v[1]:=10; v[2]:=20; v[3]:=30; WriteLn(Pick(v,2)); end."#
        ),
        &["20"]
    );
}

#[test]
fn static_array_single_element_bounds() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array[7..7] of Integer; begin a[7]:=42; WriteLn(Low(a)); WriteLn(High(a)); WriteLn(a[7]); end."#
        ),
        &["7", "7", "42"]
    );
}
