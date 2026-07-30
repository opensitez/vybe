use super::helpers::run_prints;

#[test]
fn string_intrinsic_chain_interactions_trim_len_index() {
    let out = run_prints(
        r#"
program string_intrinsic_chain_interactions_trim_len_index
    character(len=20) :: raw
    raw = '  vybe-fortran  '
    print *, len(raw)
    print *, len_trim(raw)
    print *, index(trim(raw), 'fortran')
end program string_intrinsic_chain_interactions_trim_len_index
"#,
    );
    assert_eq!(out, vec!["20", "12", "5"]);
}

#[test]
fn string_intrinsic_chain_interactions_verify_without_space() {
    let out = run_prints(
        r#"
program string_intrinsic_chain_interactions_verify_without_space
    character(len=12) :: text
    text = 'abc123def456'
    print *, verify(text, '0123456789', .true.)
    print *, scan(text, '123')
end program string_intrinsic_chain_interactions_verify_without_space
"#,
    );
    assert_eq!(out, vec!["1", "4"]);
}

#[test]
fn string_intrinsic_chain_interactions_adjustl_adjustr() {
    let out = run_prints(
        r#"
program string_intrinsic_chain_interactions_adjustl_adjustr
    character(len=10) :: text
    text = '  mix ' 
    print *, trim(adjustl(text))
    print *, len_trim(adjustr(text))
end program string_intrinsic_chain_interactions_adjustl_adjustr
"#,
    );
    assert_eq!(out, vec!["mix", "4"]);
}

#[test]
fn string_intrinsic_chain_interactions_repeat_concat_and_len() {
    let out = run_prints(
        r#"
program string_intrinsic_chain_interactions_repeat_concat_and_len
    character(len=20) :: text
    text = trim('A') // trim(repeat('B', 3))
    print *, text
    print *, len_trim(text)
    print *, index(text, 'BBB')
end program string_intrinsic_chain_interactions_repeat_concat_and_len
"#,
    );
    assert_eq!(out, vec!["ABBB", "4", "2"]);
}

#[test]
fn string_intrinsic_chain_interactions_nested_transfers() {
    let out = run_prints(
        r#"
program string_intrinsic_chain_interactions_nested_transfers
    character(len=4) :: x
    character(len=12) :: y
    x = 'ab'
    y = trim(adjustl(x // repeat('c', 2)))
    print *, y
    print *, len_trim(y)
end program string_intrinsic_chain_interactions_nested_transfers
"#,
    );
    assert_eq!(out, vec!["abcc", "4"]);
}

#[test]
fn string_intrinsic_chain_interactions_index_scan_verify_chain() {
    let out = run_prints(
        r#"
program string_intrinsic_chain_interactions_index_scan_verify_chain
    character(len=16) :: x
    x = '  one,two,three  '
    print *, index(trim(x), ',')
    print *, scan(trim(x), ',')
    print *, verify(trim(x), '1234567890, ', .true.)
end program string_intrinsic_chain_interactions_index_scan_verify_chain
"#,
    );
    assert_eq!(out, vec!["4", "4", "1"]);
}

#[test]
fn string_intrinsic_chain_interactions_prefix_suffix_chain() {
    let out = run_prints(
        r#"
program string_intrinsic_chain_interactions_prefix_suffix_chain
    character(len=12) :: text
    character(len=8) :: head
    text = '  fortran  '
    head = trim(adjustl(text))
    print *, len_trim(head)
    print *, head // '_ok'
    print *, index(head, 'tran')
end program string_intrinsic_chain_interactions_prefix_suffix_chain
"#,
    );
    assert_eq!(out, vec!["7", "fortran_ok", "4"]);
}

#[test]
fn string_intrinsic_chain_interactions_char_comparisons() {
    let out = run_prints(
        r#"
program string_intrinsic_chain_interactions_char_comparisons
    character(len=8) :: left
    character(len=8) :: right
    left = 'alpha'
    right = 'alphA'
    print *, left < right
    print *, trim(merge(left, right, left > right))
end program string_intrinsic_chain_interactions_char_comparisons
"#,
    );
    assert_eq!(out, vec!["False", "alphA"]);
}

#[test]
fn string_intrinsic_chain_interactions_replace_and_trim() {
    let out = run_prints(
        r#"
program string_intrinsic_chain_interactions_replace_and_trim
    character(len=20) :: text
    text = 'the quick brown fox'
    print *, trim(adjustl(replace(text, 'quick', 'fast')))
    print *, len_trim(adjustl(replace(text, 'quick', 'fast')))
end program string_intrinsic_chain_interactions_replace_and_trim
"#,
    );
    assert_eq!(out, vec!["the fast brown fox", "17"]);
}

#[test]
fn string_intrinsic_chain_interactions_ltrim_rtrim_adjustl_chain() {
    let out = run_prints(
        r#"
program string_intrinsic_chain_interactions_ltrim_rtrim_adjustl_chain
    character(len=12) :: text
    text = '   abc  '
    print *, len_trim(text)
    print *, len_trim(ltrim(text))
    print *, len_trim(rtrim(text))
    print *, len_trim(adjustl(text))
    print *, trim(adjustl(text))
end program string_intrinsic_chain_interactions_ltrim_rtrim_adjustl_chain
"#,
    );
    assert_eq!(out, vec!["8", "3", "6", "3", "abc"]);
}

#[test]
fn string_intrinsic_chain_interactions_scan_from_back_chain() {
    let out = run_prints(
        r#"
program string_intrinsic_chain_interactions_scan_from_back_chain
    character(len=16) :: text
    text = '  one, two, three  '
    print *, trim(adjustl(trim(text)))
    print *, scan(trim(adjustl(trim(text))), ' ,', .true.)
    print *, index(text, 'three')
end program string_intrinsic_chain_interactions_scan_from_back_chain
"#,
    );
    assert_eq!(out, vec!["one, two, three", "10", "13"]);
}

#[test]
fn string_intrinsic_chain_interactions_merge_adjustl_chain() {
    let out = run_prints(
        r#"
program string_intrinsic_chain_interactions_merge_adjustl_chain
    print *, trim(merge(adjustl('  x  '), adjustr('  y  '), len_trim('  x  ') > 0))
    print *, trim(merge(adjustl('  x  '), adjustr('  y  '), .false.))
    print *, len_trim(adjustl(merge('  xx  ', '  yy  ', .false.)))
end program string_intrinsic_chain_interactions_merge_adjustl_chain
"#,
    );
    assert_eq!(out, vec!["x", "y", "2"]);
}

#[test]
fn string_intrinsic_chain_interactions_transfer_case_chain() {
    let out = run_prints(
        r#"
program string_intrinsic_chain_interactions_transfer_case_chain
    character(len=18) :: source
    source = '  Mixed CASE Data  '
    print *, verify(adjustl(source), ' ', .false.)
    print *, adjustl(transfer(source, ''))
    print *, len_trim(transfer(trim(adjustl(source)), ''))
end program string_intrinsic_chain_interactions_transfer_case_chain
"#,
    );
    assert_eq!(out, vec!["1", "Mixed CASE Data", "16"]);
}

#[test]
fn string_intrinsic_chain_interactions_trim_scan_conditional() {
    let out = run_prints(
        r#"
program string_intrinsic_chain_interactions_trim_scan_conditional
    character(len=20) :: base
    base = 'alpha;beta;gamma'
    print *, index(trim(base), ';')
    print *, scan(trim(base), ';', .false.)
    print *, verify(merge(base, 'fallback', len(base) > 10), 'abcdefghijklmnopqrstuvwxyz', .false.)
end program string_intrinsic_chain_interactions_trim_scan_conditional
"#,
    );
    assert_eq!(out, vec!["6", "6", "6"]);
}
