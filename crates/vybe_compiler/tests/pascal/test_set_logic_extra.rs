/// Set inclusion, subset relations, and for-in iteration patterns.
use super::helpers::run_pascal;

#[test]
fn set_include_adds_member() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=(A,B,C); var s:set of TD; begin s:=[A]; Include(s,B); WriteLn(Ord(B in s)); end."#
        ),
        &["1"]
    );
}

#[test]
fn set_exclude_removes_member() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=(A,B,C); var s:set of TD; begin s:=[A,B]; Exclude(s,A); WriteLn(Ord(A in s)); end."#
        ),
        &["0"]
    );
}

#[test]
fn set_subset_strict_smaller() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=(A,B,C); var a,b:set of TD; begin a:=[A]; b:=[A,B]; WriteLn(Ord(a<=b)); end."#
        ),
        &["1"]
    );
}

#[test]
fn set_superset_strict_larger() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=(A,B,C); var a,b:set of TD; begin a:=[A,B]; b:=[A]; WriteLn(Ord(a>=b)); end."#
        ),
        &["1"]
    );
}

#[test]
fn set_equality_same_members() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=(A,B); var s,t:set of TD; begin s:=[A,B]; t:=[B,A]; WriteLn(Ord(s=t)); end."#
        ),
        &["1"]
    );
}

#[test]
fn set_union_combines() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=(A,B,C); var s,t,u:set of TD; begin s:=[A]; t:=[B]; u:=s+t; WriteLn(Ord(C in u)); WriteLn(Ord(B in u)); end."#
        ),
        &["0", "1"]
    );
}

#[test]
fn set_intersection_overlap() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=(A,B,C); var s,t,u:set of TD; begin s:=[A,B]; t:=[B,C]; u:=s*t; WriteLn(Ord(A in u)); WriteLn(Ord(B in u)); end."#
        ),
        &["0", "1"]
    );
}

#[test]
fn set_difference_removes() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=(A,B,C); var s,t,d:set of TD; begin s:=[A,B,C]; t:=[B]; d:=s-t; WriteLn(Ord(A in d)); WriteLn(Ord(B in d)); end."#
        ),
        &["1", "0"]
    );
}

#[test]
fn set_symmetric_difference_xor() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=(A,B,C); var s,t,x:set of TD; begin s:=[A,B]; t:=[B,C]; x:=s><t; WriteLn(Ord(A in x)); WriteLn(Ord(C in x)); end."#
        ),
        &["1", "1"]
    );
}

#[test]
fn for_in_enum_set_count() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=(A,B,C); var s:set of TD; c:TD; n:Integer; begin s:=[A,C]; n:=0; for c in s do Inc(n); WriteLn(n); end."#
        ),
        &["2"]
    );
}

#[test]
fn for_in_char_set_lowercase() {
    assert_eq!(
        run_pascal(
            r#"program T; var s:set of Char; c:Char; n:Integer; begin s:=['a'..'c']; n:=0; for c in s do Inc(n); WriteLn(n); end."#
        ),
        &["3"]
    );
}

#[test]
fn set_of_char_digit_range() {
    assert_eq!(
        run_pascal(
            r#"program T; var d:set of Char; begin d:=['0'..'9']; WriteLn(Ord('5' in d)); WriteLn(Ord('x' in d)); end."#
        ),
        &["1", "0"]
    );
}

#[test]
fn set_membership_not_in() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=(A,B); var s:set of TD; begin s:=[A]; WriteLn(Ord(not (B in s))); end."#
        ),
        &["1"]
    );
}

#[test]
fn set_empty_contains_nothing() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=(A,B); var s:set of TD; begin s:=[]; WriteLn(Ord(A in s)); end."#
        ),
        &["0"]
    );
}

#[test]
fn set_full_enum_all_in() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=(A,B,C); var s:set of TD; begin s:=[A..C]; WriteLn(Ord(A in s)); WriteLn(Ord(C in s)); end."#
        ),
        &["1", "1"]
    );
}

#[test]
fn set_include_loop_build() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=(A,B,C,D); var s:set of TD; i:Integer; begin s:=[]; for i:=0 to 2 do Include(s,TD(i)); WriteLn(Ord(D in s)); WriteLn(Ord(C in s)); end."#
        ),
        &["0", "1"]
    );
}

#[test]
fn set_subset_false_when_extra() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=(A,B,C); var a,b:set of TD; begin a:=[A,B]; b:=[A]; WriteLn(Ord(b<=a)); end."#
        ),
        &["1"]
    );
}

#[test]
fn set_superset_false_when_missing() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=(A,B,C); var a,b:set of TD; begin a:=[A]; b:=[A,B]; WriteLn(Ord(a>=b)); end."#
        ),
        &["0"]
    );
}

#[test]
fn for_in_empty_set_zero_iters() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=(A,B); var s:set of TD; c:TD; n:Integer; begin s:=[]; n:=0; for c in s do Inc(n); WriteLn(n); end."#
        ),
        &["0"]
    );
}

#[test]
fn set_union_with_empty_identity() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=(A,B); var s,e,u:set of TD; begin s:=[A]; e:=[]; u:=s+e; WriteLn(Ord(A in u)); end."#
        ),
        &["1"]
    );
}

#[test]
fn set_intersection_with_empty_zero() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=(A,B); var s,e,i:set of TD; begin s:=[A,B]; e:=[]; i:=s*e; WriteLn(Ord(A in i)); end."#
        ),
        &["0"]
    );
}

#[test]
fn set_char_vowel_check() {
    assert_eq!(
        run_pascal(
            r#"program T; var v:set of Char; begin v:=['a','e','i','o','u']; WriteLn(Ord('e' in v)); WriteLn(Ord('z' in v)); end."#
        ),
        &["1", "0"]
    );
}

#[test]
fn for_in_char_set_build_string() {
    assert_eq!(
        run_pascal(
            r#"program T; var s:set of Char; c:Char; t:string; begin s:=['x','y']; t:=''; for c in s do t:=t+c; WriteLn(Length(t)); end."#
        ),
        &["2"]
    );
}

#[test]
fn set_difference_self_empty() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=(A,B); var s,d:set of TD; begin s:=[A,B]; d:=s-s; WriteLn(Ord(A in d)); end."#
        ),
        &["0"]
    );
}

#[test]
fn set_assign_from_literal_brackets() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=(X,Y,Z); var s:set of TD; begin s:=[Y,Z]; WriteLn(Ord(X in s)); WriteLn(Ord(Y in s)); end."#
        ),
        &["0", "1"]
    );
}

#[test]
fn set_include_twice_idempotent() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=(A,B); var s:set of TD; n:Integer; begin s:=[]; Include(s,A); Include(s,A); n:=0; if A in s then Inc(n); WriteLn(n); end."#
        ),
        &["1"]
    );
}

#[test]
fn set_exclude_absent_no_change() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=(A,B); var s:set of TD; begin s:=[A]; Exclude(s,B); WriteLn(Ord(A in s)); end."#
        ),
        &["1"]
    );
}

#[test]
fn for_in_enum_sum_ord() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=(P,Q,R); var s:set of TD; c:TD; sum:Integer; begin s:=[P,R]; sum:=0; for c in s do sum:=sum+Ord(c); WriteLn(sum); end."#
        ),
        &["2"]
    );
}

#[test]
fn set_subset_equal_sets() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=(A,B); var s,t:set of TD; begin s:=[A,B]; t:=[A,B]; WriteLn(Ord(s<=t)); WriteLn(Ord(s>=t)); end."#
        ),
        &["1", "1"]
    );
}

#[test]
fn set_union_three_step() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=(A,B,C,D); var a,b,c,u:set of TD; begin a:=[A]; b:=[B]; c:=[C]; u:=a+b; u:=u+c; WriteLn(Ord(D in u)); WriteLn(Ord(B in u)); end."#
        ),
        &["0", "1"]
    );
}

#[test]
fn set_char_range_letters_count() {
    assert_eq!(
        run_pascal(
            r#"program T; var s:set of Char; c:Char; n:Integer; begin s:=['a'..'d']; n:=0; for c in s do Inc(n); WriteLn(n); end."#
        ),
        &["4"]
    );
}

#[test]
fn set_complement_via_full_minus() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=(A,B,C); var full,part,rest:set of TD; begin full:=[A..C]; part:=[B]; rest:=full-part; WriteLn(Ord(A in rest)); WriteLn(Ord(B in rest)); end."#
        ),
        &["1", "0"]
    );
}

#[test]
fn for_in_break_count_at_two() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=(A,B,C,D); var s:set of TD; c:TD; n:Integer; begin s:=[A..D]; n:=0; for c in s do begin Inc(n); if n=2 then Break; end; WriteLn(n); end."#
        ),
        &["2"]
    );
}

#[test]
fn set_in_record_field() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=(A,B); type TBox=record Flags:set of TD; end; var b:TBox; begin b.Flags:=[A]; Include(b.Flags,B); WriteLn(Ord(B in b.Flags)); end."#
        ),
        &["1"]
    );
}

#[test]
fn set_intersection_commutative_size() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=(A,B,C); var s,t,u,v:set of TD; c:TD; ns,nt:Integer; begin s:=[A,B]; t:=[B,C]; u:=s*t; v:=t*s; ns:=0; nt:=0; for c in u do Inc(ns); for c in v do Inc(nt); WriteLn(ns); WriteLn(nt); end."#
        ),
        &["1", "1"]
    );
}

#[test]
fn set_byte_small_range() {
    assert_eq!(
        run_pascal(
            r#"program T; var b:set of Byte; begin b:=[10..12]; WriteLn(Ord(11 in b)); WriteLn(Ord(9 in b)); end."#
        ),
        &["1", "0"]
    );
}

#[test]
fn for_in_char_set_uppercase() {
    assert_eq!(
        run_pascal(
            r#"program T; var s:set of Char; c:Char; t:Char; begin s:=['A'..'C']; t:=#0; for c in s do t:=c; WriteLn(t); end."#
        ),
        &["C"]
    );
}

#[test]
fn set_strict_subset_not_equal() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=(A,B,C); var a,b:set of TD; begin a:=[A]; b:=[A,B]; WriteLn(Ord((a<=b) and not (a=b))); end."#
        ),
        &["1"]
    );
}

#[test]
fn set_exclude_all_members() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=(A,B); var s:set of TD; begin s:=[A,B]; Exclude(s,A); Exclude(s,B); WriteLn(Ord(A in s)); end."#
        ),
        &["0"]
    );
}

#[test]
fn set_union_preserves_no_duplicates() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=(A,B,C); var s,t,u:set of TD; c:TD; n:Integer; begin s:=[A,B]; t:=[B,C]; u:=s+t; n:=0; for c in u do Inc(n); WriteLn(n); end."#
        ),
        &["3"]
    );
}

#[test]
fn for_in_singleton_set() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=(A,B,C); var s:set of TD; c:TD; v:Char; begin s:=[B]; for c in s do if c=B then v:='y' else v:='n'; WriteLn(v); end."#
        ),
        &["y"]
    );
}

#[test]
fn set_membership_in_function() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=(A,B,C); function HasB(const s:set of TD):Boolean; begin Result:=B in s; end; begin WriteLn(HasB([A,C])); WriteLn(HasB([B])); end."#
        ),
        &["false", "true"]
    );
}
