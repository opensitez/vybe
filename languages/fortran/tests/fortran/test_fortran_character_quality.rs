use super::helpers::run_prints;

#[test]
fn character_quality_trim_and_len() {
    let out = run_prints(
        r#"
program character_quality_trim_and_len
    character(len=20) :: text
    text = 'fortran   '
    print *, len(text)
    print *, len_trim(text)
    print *, len_trim(trim(text))
end program character_quality_trim_and_len
"#,
    );
    assert_eq!(out, vec!["20", "7", "7"]);
}

#[test]
fn character_quality_substring_bounds() {
    let out = run_prints(
        r#"
program character_quality_substring_bounds
    character(len=20) :: text
    text = 'runtime-check'
    print *, text(1:7)
    print *, text(9:13)
end program character_quality_substring_bounds
"#,
    );
    assert_eq!(out, vec!["runtime", "h-che"]);
}

#[test]
fn character_quality_find_substring() {
    let out = run_prints(
        r#"
program character_quality_find_substring
    character(len=20) :: source
    source = 'fortran language'
    print *, index(source, 'lang')
    print *, index(source, 'xx')
end program character_quality_find_substring
"#,
    );
    assert_eq!(out, vec!["9", "0"]);
}

#[test]
fn character_quality_index_case_sensitive() {
    let out = run_prints(
        r#"
program character_quality_index_case_sensitive
    character(len=20) :: source
    source = 'Fortran Fortran'
    print *, index(source, 'For')
    print *, index(source, 'for')
end program character_quality_index_case_sensitive
"#,
    );
    assert_eq!(out, vec!["1", "0"]);
}

#[test]
fn character_quality_verify_digits() {
    let out = run_prints(
        r#"
program character_quality_verify_digits
    character(len=20) :: source
    source = 'abc123def'
    print *, verify(source, '0123456789')
    print *, verify(source, 'abcdef')
end program character_quality_verify_digits
"#,
    );
    assert_eq!(out, vec!["1", "4"]);
}

#[test]
fn character_quality_scan_letters() {
    let out = run_prints(
        r#"
program character_quality_scan_letters
    character(len=20) :: source
    source = '0012ab34'
    print *, scan(source, 'ab')
    print *, scan(source, 'xyz')
end program character_quality_scan_letters
"#,
    );
    assert_eq!(out, vec!["5", "0"]);
}

#[test]
fn character_quality_concat_chain() {
    let out = run_prints(
        r#"
program character_quality_concat_chain
    character(len=30) :: text
    text = 'foo' // '-' // 'bar' // '-' // 'baz'
    print *, trim(text)
    print *, len_trim(text)
end program character_quality_concat_chain
"#,
    );
    assert_eq!(out, vec!["foo-bar-baz", "11"]);
}

#[test]
fn character_quality_adjust_left_right() {
    let out = run_prints(
        r#"
program character_quality_adjust_left_right
    character(len=10) :: left
    character(len=10) :: right
    left = adjustl('  hello')
    right = adjustr('hello  ')
    print *, trim(left)
    print *, trim(right)
end program character_quality_adjust_left_right
"#,
    );
    assert_eq!(out, vec!["hello", "hello"]);
}

#[test]
fn character_quality_repeat_fill() {
    let out = run_prints(
        r#"
program character_quality_repeat_fill
    character(len=20) :: text
    text = repeat('a', 5)
    print *, len_trim(text)
    print *, text
end program character_quality_repeat_fill
"#,
    );
    assert_eq!(out, vec!["5", "aaaaa"]);
}

#[test]
fn character_quality_compare_equality() {
    let out = run_prints(
        r#"
program character_quality_compare_equality
    character(len=6) :: left
    character(len=6) :: right
    left = 'abc   '
    right = 'abc   '
    if (left == right) print *, 1
    if (left /= right) print *, 0
end program character_quality_compare_equality
"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn character_quality_inequality_case() {
    let out = run_prints(
        r#"
program character_quality_inequality_case
    character(len=6) :: a
    character(len=6) :: b
    a = 'abc   '
    b = 'abd   '
    if (a < b) print *, 1
    if (a > b) print *, 0
end program character_quality_inequality_case
"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn character_quality_character_to_int_parse() {
    let out = run_prints(
        r#"
program character_quality_character_to_int_parse
    character(len=8) :: token
    integer :: value
    token = '007'
    read (token, '(I0)') value
    print *, value + 1
end program character_quality_character_to_int_parse
"#,
    );
    assert_eq!(out, vec!["8"]);
}
