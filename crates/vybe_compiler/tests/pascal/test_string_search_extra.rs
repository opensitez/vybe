/// Pos, LastDelimiter, and CompareText search/compare variants.
use super::helpers::run_pascal;

#[test]
fn pos_substring_at_start() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Pos('ab','abc')); end."#),
        &["1"]
    );
}

#[test]
fn pos_substring_at_middle() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Pos('cd','abcd')); end."#),
        &["3"]
    );
}

#[test]
fn pos_not_found_returns_zero() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Pos('z','hello')); end."#),
        &["0"]
    );
}

#[test]
fn pos_empty_needle_is_one() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Pos('','abc')); end."#),
        &["1"]
    );
}

#[test]
fn pos_single_char_in_string() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Pos('o','room')); end."#),
        &["2"]
    );
}

#[test]
fn pos_case_sensitive_miss() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Pos('A','alpha')); end."#),
        &["0"]
    );
}

#[test]
fn pos_from_offset_via_copy() {
    assert_eq!(
        run_pascal(r#"program T; var s:string; p:Integer; begin s:='banana'; p:=Pos('na',Copy(s,3,Length(s))); WriteLn(p); end."#),
        &["2"]
    );
}

#[test]
fn pos_drives_slice_extraction() {
    assert_eq!(
        run_pascal(r#"program T; var s,t:string; p:Integer; begin s:='id=42'; p:=Pos('=',s); t:=Copy(s,p+1,Length(s)); WriteLn(t); end."#),
        &["42"]
    );
}

#[test]
fn pos_overlapping_pattern() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Pos('aa','baa')); end."#),
        &["2"]
    );
}

#[test]
fn pos_at_end_position() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Pos('end','xxend')); end."#),
        &["3"]
    );
}

#[test]
fn lastdelimiter_comma_path() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(LastDelimiter('/\', 'dir/file.txt')); end."#),
        &["4"]
    );
}

#[test]
fn lastdelimiter_backslash_windows() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(LastDelimiter('/\', 'C:\tmp\a')); end."#),
        &["7"]
    );
}

#[test]
fn lastdelimiter_none_returns_zero() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(LastDelimiter(',;', 'plain')); end."#),
        &["0"]
    );
}

#[test]
fn lastdelimiter_multiple_chars_set() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(LastDelimiter(',;', 'a,b;c')); end."#),
        &["4"]
    );
}

#[test]
fn lastdelimiter_trailing_delim() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(LastDelimiter('/', 'root/sub/')); end."#),
        &["9"]
    );
}

#[test]
fn lastdelimiter_extract_filename() {
    assert_eq!(
        run_pascal(r#"program T; var p:Integer; s:string; begin s:='path/to/name'; p:=LastDelimiter('/',s); WriteLn(Copy(s,p+1,Length(s))); end."#),
        &["name"]
    );
}

#[test]
fn lastdelimiter_pipe_list() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(LastDelimiter('|', 'a|b|c')); end."#),
        &["3"]
    );
}

#[test]
fn lastdelimiter_space_words() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(LastDelimiter(' ', 'one two three')); end."#),
        &["8"]
    );
}

#[test]
fn comparetext_equal_ignore_case() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(CompareText('Hello','hello')); end."#),
        &["0"]
    );
}

#[test]
fn comparetext_less_upper_before_lower() {
    assert_eq!(
        run_pascal(r#"program T; var r:Integer; begin r:=CompareText('apple','Banana'); if r<0 then WriteLn('less') else WriteLn('geq'); end."#),
        &["less"]
    );
}

#[test]
fn comparetext_greater_longer_prefix() {
    assert_eq!(
        run_pascal(r#"program T; var r:Integer; begin r:=CompareText('zzz','zz'); if r>0 then WriteLn('gt') else WriteLn('le'); end."#),
        &["gt"]
    );
}

#[test]
fn comparetext_empty_vs_empty() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(CompareText('','')); end."#),
        &["0"]
    );
}

#[test]
fn comparetext_empty_vs_nonempty() {
    assert_eq!(
        run_pascal(r#"program T; var r:Integer; begin r:=CompareText('','a'); if r<0 then WriteLn('less') else WriteLn('geq'); end."#),
        &["less"]
    );
}

#[test]
fn comparetext_numbers_as_strings() {
    assert_eq!(
        run_pascal(r#"program T; var r:Integer; begin r:=CompareText('10','2'); if r<0 then WriteLn('less') else WriteLn('geq'); end."#),
        &["less"]
    );
}

#[test]
fn comparetext_mixed_case_order() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(CompareText('AbC','aBc')); end."#),
        &["0"]
    );
}

#[test]
fn comparetext_sort_three_words() {
    assert_eq!(
        run_pascal(r#"program T; function Before(const a,b:string):Boolean; begin Result:=CompareText(a,b)<0; end; begin if Before('ant','Bee') then WriteLn('ok'); end."#),
        &["ok"]
    );
}

#[test]
fn comparetext_underscore_vs_letter() {
    assert_eq!(
        run_pascal(r#"program T; var r:Integer; begin r:=CompareText('_a','a'); if r<0 then WriteLn('u') else WriteLn('v'); end."#),
        &["u"]
    );
}

#[test]
fn comparetext_same_length_diff_char() {
    assert_eq!(
        run_pascal(r#"program T; var r:Integer; begin r:=CompareText('cat','car'); if r>0 then WriteLn('t>r') else WriteLn('other'); end."#),
        &["t>r"]
    );
}

#[test]
fn pos_then_comparetext_equal_tail() {
    assert_eq!(
        run_pascal(r#"program T; var s,t:string; p:Integer; begin s:='pre-TAIL'; p:=Pos('TAIL',s); t:=Copy(s,p,4); if CompareText(t,'tail')=0 then WriteLn('match'); end."#),
        &["match"]
    );
}

#[test]
fn lastdelimiter_then_pos_rejoin() {
    assert_eq!(
        run_pascal(r#"program T; var s:string; d,p:Integer; begin s:='a.b.c'; d:=LastDelimiter('.',s); p:=Pos('.',s); WriteLn(d-p); end."#),
        &["2"]
    );
}

#[test]
fn pos_in_empty_haystack_zero() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Pos('a','')); end."#),
        &["0"]
    );
}

#[test]
fn pos_full_string_match() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Pos('same','same')); end."#),
        &["1"]
    );
}

#[test]
fn comparetext_single_char_pair() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(CompareText('x','X')); end."#),
        &["0"]
    );
}

#[test]
fn lastdelimiter_at_first_char() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(LastDelimiter('/', '/root')); end."#),
        &["1"]
    );
}

#[test]
fn pos_loop_find_all_occurrences_count() {
    assert_eq!(
        run_pascal(r#"program T; var s:string; i,p,c:Integer; begin s:='abab'; c:=0; i:=1; repeat p:=Pos('a',Copy(s,i,Length(s))); if p>0 then begin Inc(c); i:=i+p; end; until p=0; WriteLn(c); end."#),
        &["2"]
    );
}

#[test]
fn comparetext_symmetry_check() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(CompareText('foo','bar')); WriteLn(CompareText('bar','foo')); end."#),
        &["1", "-1"]
    );
}

#[test]
fn lastdelimiter_colon_drive() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(LastDelimiter(':', 'C:folder')); end."#),
        &["2"]
    );
}

#[test]
fn pos_after_delim_extension() {
    assert_eq!(
        run_pascal(r#"program T; var s:string; d,p:Integer; begin s:='file.txt'; d:=LastDelimiter('.',s); p:=Pos('.',s); WriteLn(Copy(s,d+1,3)); end."#),
        &["txt"]
    );
}

#[test]
fn comparetext_prefix_relation() {
    assert_eq!(
        run_pascal(r#"program T; var r:Integer; begin r:=CompareText('test','testing'); if r<0 then WriteLn('shorter') else WriteLn('other'); end."#),
        &["shorter"]
    );
}

#[test]
fn pos_multichar_second_occurrence_manual() {
    assert_eq!(
        run_pascal(r#"program T; var s:string; p1,p2:Integer; begin s:='xyxy'; p1:=Pos('xy',s); p2:=Pos('xy',Copy(s,p1+1,Length(s))); WriteLn(p1); WriteLn(p2); end."#),
        &["1", "2"]
    );
}

#[test]
fn lastdelimiter_semicolon_csv() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(LastDelimiter(',', '1,2,3,4')); end."#),
        &["5"]
    );
}

#[test]
fn comparetext_with_trimmed_input() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(CompareText(Trim('  Hi '),'hi')); end."#),
        &["0"]
    );
}
