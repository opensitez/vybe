/// Scoped enums, set operations, and for-in iteration on sets.
use super::helpers::run_pascal;

#[test]
fn scoped_enum_assign_and_ord() {
    assert_eq!(
        run_pascal(
            r#"program T; type TLevel=(Low, Mid, High); var l:TLevel; begin l:=High; WriteLn(Ord(l)); end."#
        ),
        &["2"]
    );
}

#[test]
fn scoped_enum_explicit_value_gap() {
    assert_eq!(
        run_pascal(
            r#"program T; type TCode=(Ok=0, Warn=5, Fail=10); var c:TCode; begin c:=Fail; WriteLn(Ord(c)); end."#
        ),
        &["10"]
    );
}

#[test]
fn set_union_two_enum_members() {
    assert_eq!(
        run_pascal(
            r#"program T; type TF=(A,B,C); var s1,s2,s: set of TF; begin s1:=[A]; s2:=[B]; s:=s1+s2; WriteLn(Ord(A in s)); WriteLn(Ord(B in s)); end."#
        ),
        &["1", "1"]
    );
}

#[test]
fn set_intersection_overlapping_members() {
    assert_eq!(
        run_pascal(
            r#"program T; type TF=(A,B,C); var s1,s2,s: set of TF; begin s1:=[A,B]; s2:=[B,C]; s:=s1*s2; WriteLn(B in s); WriteLn(A in s); end."#
        ),
        &["true", "false"]
    );
}

#[test]
fn set_difference_removes_member() {
    assert_eq!(
        run_pascal(
            r#"program T; type TF=(A,B,C); var s1,s2,s: set of TF; begin s1:=[A,B,C]; s2:=[B]; s:=s1-s2; WriteLn(A in s); WriteLn(B in s); end."#
        ),
        &["true", "false"]
    );
}

#[test]
fn set_symmetric_difference_xor() {
    assert_eq!(
        run_pascal(
            r#"program T; type TF=(A,B); var s1,s2,s: set of TF; begin s1:=[A]; s2:=[B]; s:=s1><s2; WriteLn(A in s); WriteLn(B in s); end."#
        ),
        &["true", "true"]
    );
}

#[test]
fn set_subset_superset_relations() {
    assert_eq!(
        run_pascal(
            r#"program T; type TF=(A,B,C); var small,big: set of TF; begin small:=[A]; big:=[A,B]; WriteLn(small<=big); WriteLn(big>=small); end."#
        ),
        &["TRUE", "TRUE"]
    );
}

#[test]
fn set_include_adds_member() {
    assert_eq!(
        run_pascal(
            r#"program T; type TF=(X,Y); var s: set of TF; begin s:=[X]; Include(s,Y); WriteLn(Y in s); end."#
        ),
        &["true"]
    );
}

#[test]
fn set_exclude_removes_member() {
    assert_eq!(
        run_pascal(
            r#"program T; type TF=(X,Y); var s: set of TF; begin s:=[X,Y]; Exclude(s,X); WriteLn(X in s); WriteLn(Y in s); end."#
        ),
        &["false", "true"]
    );
}

#[test]
fn for_in_over_enum_array_literal() {
    assert_eq!(
        run_pascal(
            r#"program T; type TDay=(Mon,Tue,Wed); var d:TDay; n:Integer; begin n:=0; for d in [Mon,Tue,Wed] do Inc(n); WriteLn(n); end."#
        ),
        &["3"]
    );
}

#[test]
fn for_in_over_subset_enum_values() {
    assert_eq!(
        run_pascal(
            r#"program T; type TC=(R,G,B); var c:TC; s:String; begin s:=''; for c in [R,B] do s:=s+'x'; WriteLn(Length(s)); end."#
        ),
        &["2"]
    );
}

#[test]
fn set_of_char_range_membership() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD='0'..'9'; var s: set of TD; begin s:=['1','3','5']; WriteLn('3' in s); WriteLn('2' in s); end."#
        ),
        &["true", "false"]
    );
}

#[test]
fn set_empty_equals_empty() {
    assert_eq!(
        run_pascal(r#"program T; type TF=(A,B); var s1,s2: set of TF; begin WriteLn(s1=s2); end."#),
        &["TRUE"]
    );
}

#[test]
fn enum_case_with_scoped_type() {
    assert_eq!(
        run_pascal(
            r#"program T; type TS=(Off,On); var s:TS; begin s:=On; case s of Off:WriteLn('0'); On:WriteLn('1'); end; end."#
        ),
        &["1"]
    );
}

#[test]
fn set_union_with_empty_is_identity() {
    assert_eq!(
        run_pascal(
            r#"program T; type TF=(A); var s,e,u: set of TF; begin s:=[A]; u:=s+e; WriteLn(A in u); end."#
        ),
        &["true"]
    );
}

#[test]
fn set_intersection_with_empty_is_empty() {
    assert_eq!(
        run_pascal(
            r#"program T; type TF=(A,B); var s,e,i: set of TF; begin s:=[A,B]; i:=s*e; WriteLn(Length(i)); end."#
        ),
        &["0"]
    );
}

#[test]
fn enum_succ_from_first_to_second() {
    assert_eq!(
        run_pascal(
            r#"program T; type TE=(Alpha,Beta,Gamma); var e:TE; begin e:=Alpha; e:=Succ(e); WriteLn(Ord(e)); end."#
        ),
        &["1"]
    );
}

#[test]
fn enum_pred_from_last_to_middle() {
    assert_eq!(
        run_pascal(
            r#"program T; type TE=(Alpha,Beta,Gamma); var e:TE; begin e:=Gamma; e:=Pred(e); WriteLn(Ord(e)); end."#
        ),
        &["1"]
    );
}

#[test]
fn set_in_record_field_storage() {
    assert_eq!(
        run_pascal(
            r#"program T; type TF=(Read,Write); type TMode=record Flags: set of TF; end; var m:TMode; begin m.Flags:=[Read,Write]; WriteLn(Write in m.Flags); end."#
        ),
        &["true"]
    );
}

#[test]
fn for_in_enum_accumulate_ord_sum() {
    assert_eq!(
        run_pascal(
            r#"program T; type TN=(N0,N1,N2); var n:TN; sum:Integer; begin sum:=0; for n in [N0,N1,N2] do sum:=sum+Ord(n); WriteLn(sum); end."#
        ),
        &["3"]
    );
}

#[test]
fn set_complement_via_full_minus() {
    assert_eq!(
        run_pascal(
            r#"program T; type TF=(A,B,C); var full,part,rest: set of TF; begin full:=[A,B,C]; part:=[B]; rest:=full-part; WriteLn(A in rest); WriteLn(C in rest); end."#
        ),
        &["true", "true"]
    );
}

#[test]
fn scoped_enum_function_param() {
    assert_eq!(
        run_pascal(
            r#"program T; type TMode=(Safe,Fast); function Tag(m:TMode):Integer; begin Result:=Ord(m); end; begin WriteLn(Tag(Fast)); end."#
        ),
        &["1"]
    );
}

#[test]
fn set_of_byte_range_literals() {
    assert_eq!(
        run_pascal(r#"program T; var s: set of Byte; begin s:=[1,2,3]; WriteLn(2 in s); end."#),
        &["true"]
    );
}

#[test]
fn enum_array_index_by_ord() {
    assert_eq!(
        run_pascal(
            r#"program T; type TC=(Red,Green,Blue); var names:array[TC] of String; begin names[Red]:='r'; names[Blue]:='b'; WriteLn(names[Blue]); end."#
        ),
        &["b"]
    );
}

#[test]
fn set_membership_negated_not_in() {
    assert_eq!(
        run_pascal(
            r#"program T; type TF=(A,B); var s: set of TF; begin s:=[A]; WriteLn(not (B in s)); end."#
        ),
        &["true"]
    );
}

#[test]
fn for_in_char_set_literal() {
    assert_eq!(
        run_pascal(
            r#"program T; var ch:Char; n:Integer; begin n:=0; for ch in ['a','b'] do Inc(n); WriteLn(n); end."#
        ),
        &["2"]
    );
}

#[test]
fn enum_compare_equality() {
    assert_eq!(
        run_pascal(
            r#"program T; type TS=(On,Off); var a,b:TS; begin a:=On; b:=On; WriteLn(a=b); end."#
        ),
        &["TRUE"]
    );
}

#[test]
fn set_assign_from_literal_brackets() {
    assert_eq!(
        run_pascal(
            r#"program T; type TF=(P,Q,R); var s: set of TF; begin s:=[P,R]; WriteLn(P in s); WriteLn(Q in s); end."#
        ),
        &["true", "false"]
    );
}

#[test]
fn scoped_enum_in_variant_record_tag() {
    assert_eq!(
        run_pascal(
            r#"program T; type TK=(IntK,StrK); type TV=record case TK of IntK:(I:Integer); StrK:(S:String); end; var v:TV; begin v.I:=7; WriteLn(v.I); end."#
        ),
        &["7"]
    );
}

#[test]
fn set_union_three_way_chain() {
    assert_eq!(
        run_pascal(
            r#"program T; type TF=(A,B,C,D); var s: set of TF; begin s:=[A]+[B]+[C]; WriteLn(Length(s)); end."#
        ),
        &["3"]
    );
}

#[test]
fn enum_for_in_breaks_on_count() {
    assert_eq!(
        run_pascal(
            r#"program T; type TN=(N1,N2,N3,N4); var n:TN; c:Integer; begin c:=0; for n in [N1,N2,N3] do Inc(c); WriteLn(c); end."#
        ),
        &["3"]
    );
}

#[test]
fn set_difference_self_is_empty() {
    assert_eq!(
        run_pascal(
            r#"program T; type TF=(A,B); var s,d: set of TF; begin s:=[A,B]; d:=s-s; WriteLn(Length(d)); end."#
        ),
        &["0"]
    );
}

#[test]
fn enum_explicit_start_at_ten() {
    assert_eq!(
        run_pascal(
            r#"program T; type TErr=(None=0, Minor=10, Major=20); var e:TErr; begin e:=Minor; WriteLn(Ord(e)); end."#
        ),
        &["10"]
    );
}

#[test]
fn set_superset_strict_false_when_equal() {
    assert_eq!(
        run_pascal(r#"program T; type TF=(A); var s: set of TF; begin s:=[A]; WriteLn(s>s); end."#),
        &["FALSE"]
    );
}

#[test]
fn for_in_empty_enum_array_zero_iterations() {
    assert_eq!(
        run_pascal(
            r#"program T; type TF=(A,B); var f:TF; n:Integer; begin n:=0; for f in [] do Inc(n); WriteLn(n); end."#
        ),
        &["0"]
    );
}

#[test]
fn set_char_add_via_include_loop() {
    assert_eq!(
        run_pascal(
            r#"program T; var s: set of Char; begin Include(s,'a'); Include(s,'b'); WriteLn(Length(s)); end."#
        ),
        &["2"]
    );
}

#[test]
fn enum_record_array_of_values() {
    assert_eq!(
        run_pascal(
            r#"program T; type TC=(Cold,Warm,Hot); var temps:array[0..2] of TC; begin temps[0]:=Cold; temps[2]:=Hot; WriteLn(Ord(temps[2])); end."#
        ),
        &["2"]
    );
}

#[test]
fn set_intersection_result_is_subset_of_both() {
    assert_eq!(
        run_pascal(
            r#"program T; type TF=(A,B,C); var s1,s2,i: set of TF; begin s1:=[A,B]; s2:=[B,C]; i:=s1*s2; WriteLn(i<=s1); WriteLn(i<=s2); end."#
        ),
        &["TRUE", "TRUE"]
    );
}

#[test]
fn scoped_enum_return_from_function() {
    assert_eq!(
        run_pascal(
            r#"program T; type TDir=(North,East,South,West); function DefaultDir:TDir; begin Result:=North; end; begin WriteLn(Ord(DefaultDir)); end."#
        ),
        &["0"]
    );
}

#[test]
fn set_equality_after_include() {
    assert_eq!(
        run_pascal(
            r#"program T; type TF=(A,B); var s1,s2: set of TF; begin s1:=[A]; s2:=[A]; Include(s2,B); WriteLn(s1=s2); end."#
        ),
        &["FALSE"]
    );
}
