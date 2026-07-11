/// For-in iterator loops over arrays, strings, sets, and enums.
use super::helpers::run_pascal;

#[test]
fn for_in_static_integer_array_sum() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array[0..3] of Integer; x, s: Integer; begin a[0]:=1; a[1]:=2; a[2]:=3; a[3]:=4; s:=0; for x in a do s:=s+x; WriteLn(s); end."#
        ),
        &["10"]
    );
}

#[test]
fn for_in_dynamic_array_literal() {
    assert_eq!(
        run_pascal(r#"program T; var x: Integer; begin for x in [10,20,30] do WriteLn(x); end."#),
        &["10", "20", "30"]
    );
}

#[test]
fn for_in_string_chars_uppercase() {
    assert_eq!(
        run_pascal(r#"program T; var c: Char; begin for c in 'ab' do WriteLn(UpCase(c)); end."#),
        &["A", "B"]
    );
}

#[test]
fn for_in_set_members_print_ord() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=(One,Two,Three); var d: TD; begin for d in [One,Three] do WriteLn(Ord(d)); end."#
        ),
        &["0", "2"]
    );
}

#[test]
fn for_in_array_break_on_target() {
    assert_eq!(
        run_pascal(
            r#"program T; var x: Integer; begin for x in [1,2,3,4,5] do begin if x=3 then break; WriteLn(x); end; end."#
        ),
        &["1", "2"]
    );
}

#[test]
fn for_in_array_continue_skip_evens() {
    assert_eq!(
        run_pascal(
            r#"program T; var x, s: Integer; begin s:=0; for x in [1,2,3,4] do begin if x mod 2=0 then continue; s:=s+x; end; WriteLn(s); end."#
        ),
        &["4"]
    );
}

#[test]
fn for_in_nested_outer_inner() {
    assert_eq!(
        run_pascal(
            r#"program T; var i,j,c: Integer; begin c:=0; for i in [1,2] do for j in [1,2] do c:=c+1; WriteLn(c); end."#
        ),
        &["4"]
    );
}

#[test]
fn for_in_string_build_word() {
    assert_eq!(
        run_pascal(
            r#"program T; var c: Char; s: string; begin s:=''; for c in 'hi' do s:=s+c; WriteLn(s); end."#
        ),
        &["hi"]
    );
}

#[test]
fn for_in_real_dynamic_array() {
    assert_eq!(
        run_pascal(
            r#"program T; var r: Double; s: Double; begin s:=0; for r in [1.5, 2.5] do s:=s+r; WriteLn(Round(s)); end."#
        ),
        &["4"]
    );
}

#[test]
fn for_in_boolean_flags_count_true() {
    assert_eq!(
        run_pascal(
            r#"program T; var b: Boolean; c: Integer; begin c:=0; for b in [true,false,true] do if b then c:=c+1; WriteLn(c); end."#
        ),
        &["2"]
    );
}

#[test]
fn for_in_record_array_field_access() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; end; var items: array[0..1] of TR; it: TR; s: Integer; begin items[0].V:=2; items[1].V:=3; s:=0; for it in items do s:=s+it.V; WriteLn(s); end."#
        ),
        &["5"]
    );
}

#[test]
fn for_in_enum_weekdays_count() {
    assert_eq!(
        run_pascal(
            r#"program T; type TDay=(Mon,Tue,Wed,Thu,Fri); var d: TDay; n: Integer; begin n:=0; for d in [Mon,Tue,Wed,Thu,Fri] do n:=n+1; WriteLn(n); end."#
        ),
        &["5"]
    );
}

#[test]
fn for_in_string_count_vowels() {
    assert_eq!(
        run_pascal(
            r#"program T; var c: Char; n: Integer; begin n:=0; for c in 'aeiou' do n:=n+1; WriteLn(n); end."#
        ),
        &["5"]
    );
}

#[test]
fn for_in_array_find_first_negative() {
    assert_eq!(
        run_pascal(
            r#"program T; var x, hit: Integer; begin hit:=0; for x in [3,5,-1,9] do if x<0 then begin hit:=x; break; end; WriteLn(hit); end."#
        ),
        &["-1"]
    );
}

#[test]
fn for_in_procedure_param_array() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure ShowAll(a: array of Integer); var x: Integer; begin for x in a do WriteLn(x); end; begin ShowAll([7,8]); end."#
        ),
        &["7", "8"]
    );
}

#[test]
fn for_in_char_set_membership_build() {
    assert_eq!(
        run_pascal(
            r#"program T; var c: Char; s: set of Char; begin s:=[]; for c in 'abc' do Include(s,c); if 'b' in s then WriteLn('ok'); end."#
        ),
        &["ok"]
    );
}

#[test]
fn for_in_empty_dynamic_array_runs_zero_times() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array of Integer; x, n: Integer; begin a:=[]; n:=0; for x in a do n:=n+1; WriteLn(n); end."#
        ),
        &["0"]
    );
}

#[test]
fn for_in_single_element_array() {
    assert_eq!(
        run_pascal(r#"program T; var x: Integer; begin for x in [42] do WriteLn(x); end."#),
        &["42"]
    );
}

#[test]
fn for_in_negative_integer_array() {
    assert_eq!(
        run_pascal(
            r#"program T; var x, m: Integer; begin m:=0; for x in [-3,-1,2] do if x>m then m:=x; WriteLn(m); end."#
        ),
        &["2"]
    );
}

#[test]
fn for_in_string_with_spaces_trim_check() {
    assert_eq!(
        run_pascal(
            r#"program T; var c: Char; hasSpace: Boolean; begin hasSpace:=false; for c in 'a b' do if c=' ' then hasSpace:=true; if hasSpace then WriteLn('space'); end."#
        ),
        &["space"]
    );
}

#[test]
fn for_in_multiline_string_chars() {
    assert_eq!(
        run_pascal(
            r#"program T; var c: Char; n: Integer; begin n:=0; for c in 'xy' do n:=n+1; WriteLn(n); end."#
        ),
        &["2"]
    );
}

#[test]
fn for_in_array_product() {
    assert_eq!(
        run_pascal(
            r#"program T; var x, p: Integer; begin p:=1; for x in [1,2,3] do p:=p*x; WriteLn(p); end."#
        ),
        &["6"]
    );
}

#[test]
fn for_in_class_method_array_param() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBox=class class procedure Sum(a: array of Integer); var x,s: Integer; begin s:=0; for x in a do s:=s+x; WriteLn(s); end; end; begin TBox.Sum([4,5,6]); end."#
        ),
        &["15"]
    );
}

#[test]
fn for_in_subrange_enum_slice() {
    assert_eq!(
        run_pascal(
            r#"program T; type TColor=(Red,Green,Blue,Yellow); var c: TColor; n: Integer; begin n:=0; for c in [Green,Blue] do n:=n+1; WriteLn(n); end."#
        ),
        &["2"]
    );
}

#[test]
fn for_in_string_index_not_used_but_runs() {
    assert_eq!(
        run_pascal(
            r#"program T; var c: Char; s: string; begin s:=''; for c in 'Z' do s:=c+s; WriteLn(s); end."#
        ),
        &["Z"]
    );
}

#[test]
fn for_in_integer_set_literal_sum() {
    assert_eq!(
        run_pascal(
            r#"program T; var x, t: Integer; begin t:=0; for x in [100,200] do t:=t div 100; WriteLn(t); end."#
        ),
        &["2"]
    );
}
