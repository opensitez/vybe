/// Extended dynamic arrays: SetLength, slice, Copy, concat patterns.
use super::helpers::run_pascal;

#[test]
fn setlength_from_zero_to_one() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of Integer; begin SetLength(a,1); a[0]:=9; WriteLn(a[0]); end."#
        ),
        &["9"]
    );
}

#[test]
fn setlength_double_then_shrink() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of Integer; begin SetLength(a,4); SetLength(a,2); WriteLn(Length(a)); end."#
        ),
        &["2"]
    );
}

#[test]
fn copy_dynamic_slice_middle_two() {
    assert_eq!(
        run_pascal(
            r#"program T; var src,dst:array of Integer; begin src:=[10,20,30,40]; dst:=Copy(src,2,2); WriteLn(dst[0]); WriteLn(dst[1]); end."#
        ),
        &["20", "30"]
    );
}

#[test]
fn copy_from_start_single_element() {
    assert_eq!(
        run_pascal(
            r#"program T; var src,dst:array of Integer; begin src:=[5,6,7]; dst:=Copy(src,1,1); WriteLn(dst[0]); end."#
        ),
        &["5"]
    );
}

#[test]
fn concat_two_dynamic_integer_arrays() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b,c:array of Integer; begin a:=[1,2]; b:=[3]; SetLength(c,Length(a)+Length(b)); c[0]:=a[0]; c[1]:=a[1]; c[2]:=b[0]; WriteLn(c[2]); end."#
        ),
        &["3"]
    );
}

#[test]
fn setlength_preserves_first_element_on_grow() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of Integer; begin SetLength(a,1); a[0]:=77; SetLength(a,3); WriteLn(a[0]); end."#
        ),
        &["77"]
    );
}

#[test]
fn dynamic_array_assign_copies_values() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:array of Integer; begin a:=[1,2]; b:=a; b[0]:=9; WriteLn(a[0]); WriteLn(b[0]); end."#
        ),
        &["1", "9"]
    );
}

#[test]
fn setlength_string_array_two_items() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of String; begin SetLength(a,2); a[0]:='foo'; a[1]:='bar'; WriteLn(a[1]); end."#
        ),
        &["bar"]
    );
}

#[test]
fn high_after_setlength_five() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of Integer; begin SetLength(a,5); WriteLn(High(a)); end."#
        ),
        &["4"]
    );
}

#[test]
fn low_always_zero_dynamic() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of Integer; begin SetLength(a,2); WriteLn(Low(a)); end."#
        ),
        &["0"]
    );
}

#[test]
fn iterate_dynamic_with_for_to_high() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of Integer; i:Integer; begin a:=[2,4,6]; for i:=0 to High(a) do WriteLn(a[i]); end."#
        ),
        &["2", "4", "6"]
    );
}

#[test]
fn setlength_zero_then_grow_again() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of Integer; begin a:=[1]; SetLength(a,0); SetLength(a,1); a[0]:=5; WriteLn(a[0]); end."#
        ),
        &["5"]
    );
}

#[test]
fn copy_string_array_fragment() {
    assert_eq!(
        run_pascal(
            r#"program T; var src,dst:array of String; begin src:=['a','b','c']; dst:=Copy(src,2,2); WriteLn(dst[1]); end."#
        ),
        &["c"]
    );
}

#[test]
fn dynamic_array_in_function_sum() {
    assert_eq!(
        run_pascal(
            r#"program T; function Total(const a:array of Integer):Integer; var i:Integer; begin Result:=0; for i:=Low(a) to High(a) do Result:=Result+a[i]; end; var d:array of Integer; begin d:=[1,2,3]; WriteLn(Total(d)); end."#
        ),
        &["6"]
    );
}

#[test]
fn setlength_triple_growth_chain() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of Integer; begin SetLength(a,1); SetLength(a,2); SetLength(a,3); WriteLn(Length(a)); end."#
        ),
        &["3"]
    );
}

#[test]
fn concat_via_literal_plus_dynamic() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:array of Integer; begin a:=[1]; b:=[2,3]; a:=a+b; WriteLn(Length(a)); WriteLn(a[2]); end."#
        ),
        &["3", "3"]
    );
}

#[test]
fn copy_full_array_same_length() {
    assert_eq!(
        run_pascal(
            r#"program T; var src,dst:array of Integer; begin src:=[4,5]; dst:=Copy(src,1,2); WriteLn(dst[1]); end."#
        ),
        &["5"]
    );
}

#[test]
fn setlength_on_record_dynamic_field() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB=record Items:array of Integer; end; var b:TB; begin SetLength(b.Items,2); b.Items[1]:=12; WriteLn(b.Items[1]); end."#
        ),
        &["12"]
    );
}

#[test]
fn dynamic_array_boolean_all_true() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:array of Boolean; i:Integer; ok:Boolean; begin f:=[true,true]; ok:=true; for i:=0 to High(f) do if not f[i] then ok:=false; WriteLn(ok); end."#
        ),
        &["true"]
    );
}

#[test]
fn slice_last_element_via_index() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of Integer; begin a:=[1,2,3,4]; WriteLn(a[High(a)]); end."#
        ),
        &["4"]
    );
}

#[test]
fn setlength_char_dynamic_array() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of Char; begin SetLength(a,3); a[0]:='a'; a[1]:='b'; a[2]:='c'; WriteLn(a[1]); end."#
        ),
        &["b"]
    );
}

#[test]
fn copy_beyond_length_clamped_behavior() {
    assert_eq!(
        run_pascal(
            r#"program T; var src,dst:array of Integer; begin src:=[1,2]; dst:=Copy(src,2,5); WriteLn(Length(dst)); end."#
        ),
        &["1"]
    );
}

#[test]
fn concat_empty_with_nonempty() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:array of Integer; begin SetLength(a,0); b:=[7]; a:=a+b; WriteLn(a[0]); end."#
        ),
        &["7"]
    );
}

#[test]
fn dynamic_array_reverse_in_place() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of Integer; t:Integer; begin a:=[1,2,3]; t:=a[0]; a[0]:=a[2]; a[2]:=t; WriteLn(a[0]); WriteLn(a[2]); end."#
        ),
        &["3", "1"]
    );
}

#[test]
fn setlength_real_array_fill() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of Double; i:Integer; begin SetLength(a,3); for i:=0 to 2 do a[i]:=1.0; WriteLn(Round(a[2])); end."#
        ),
        &["1"]
    );
}

#[test]
fn function_returns_copy_of_input() {
    assert_eq!(
        run_pascal(
            r#"program T; function Clone(const a:array of Integer):array of Integer; begin Result:=Copy(a,1,Length(a)); end; var x,y:array of Integer; begin x:=[8,9]; y:=Clone(x); WriteLn(y[1]); end."#
        ),
        &["9"]
    );
}

#[test]
fn dynamic_array_nested_setlength_rows() {
    assert_eq!(
        run_pascal(
            r#"program T; var r1,r2:array of Integer; begin SetLength(r1,2); SetLength(r2,3); r2[2]:=5; WriteLn(r2[2]); end."#
        ),
        &["5"]
    );
}

#[test]
fn concat_three_small_arrays_manual() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b,c,d:array of Integer; begin a:=[1]; b:=[2]; c:=[3]; d:=a+b+c; WriteLn(d[2]); end."#
        ),
        &["3"]
    );
}

#[test]
fn setlength_shrink_keeps_low_indices() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of Integer; begin SetLength(a,4); a[0]:=1; a[1]:=2; SetLength(a,2); WriteLn(a[0]); WriteLn(a[1]); end."#
        ),
        &["1", "2"]
    );
}

#[test]
fn copy_single_char_from_string_as_array() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of Char; begin a:=['x','y','z']; WriteLn(a[1]); end."#
        ),
        &["y"]
    );
}

#[test]
fn dynamic_array_param_high_bound() {
    assert_eq!(
        run_pascal(
            r#"program T; function LastIdx(const a:array of Integer):Integer; begin Result:=High(a); end; var d:array of Integer; begin d:=[0,0,0]; WriteLn(LastIdx(d)); end."#
        ),
        &["2"]
    );
}

#[test]
fn setlength_append_slot_pattern() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of Integer; n:Integer; begin a:=[1,2]; n:=Length(a); SetLength(a,n+1); a[n]:=3; WriteLn(a[2]); end."#
        ),
        &["3"]
    );
}

#[test]
fn slice_copy_first_only() {
    assert_eq!(
        run_pascal(
            r#"program T; var src,dst:array of Integer; begin src:=[9,8,7]; dst:=Copy(src,1,1); WriteLn(dst[0]); end."#
        ),
        &["9"]
    );
}

#[test]
fn dynamic_array_clear_by_setlength_zero() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of Integer; begin a:=[1,2,3]; SetLength(a,0); WriteLn(Length(a)); end."#
        ),
        &["0"]
    );
}

#[test]
fn concat_string_arrays() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:array of String; begin a:=['hi']; b:=['there']; a:=a+b; WriteLn(a[1]); end."#
        ),
        &["there"]
    );
}

#[test]
fn setlength_enum_array() {
    assert_eq!(
        run_pascal(
            r#"program T; type TC=(Red,Green); var a:array of TC; begin SetLength(a,2); a[1]:=Green; WriteLn(Ord(a[1])); end."#
        ),
        &["1"]
    );
}

#[test]
fn copy_empty_when_count_zero() {
    assert_eq!(
        run_pascal(
            r#"program T; var src,dst:array of Integer; begin src:=[1,2,3]; dst:=Copy(src,2,0); WriteLn(Length(dst)); end."#
        ),
        &["0"]
    );
}

#[test]
fn dynamic_array_var_param_reverse() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Rev(var a:array of Integer); var t:Integer; begin t:=a[0]; a[0]:=a[1]; a[1]:=t; end; var d:array of Integer; begin d:=[3,4]; Rev(d); WriteLn(d[0]); end."#
        ),
        &["4"]
    );
}

#[test]
fn setlength_large_twenty_elements() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of Integer; begin SetLength(a,20); a[19]:=42; WriteLn(a[19]); end."#
        ),
        &["42"]
    );
}

#[test]
fn concat_self_with_literal_one() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of Integer; begin a:=[1]; a:=a+[2]; WriteLn(Length(a)); WriteLn(a[1]); end."#
        ),
        &["2", "2"]
    );
}
