use super::helpers::run_prints;

#[test]
fn test_strspn_positive_offset() {
    assert_eq!(
        run_prints(
            r#"<?php
echo strspn('foo123bar', '0123456789', 3), "\n";
"#
        ),
        vec!["3"]
    );
}

#[test]
fn test_strspn_offset_and_max_length() {
    assert_eq!(
        run_prints(
            r#"<?php
echo strspn('foo123456bar', '0123456789', 3, 2), "\n";
"#
        ),
        vec!["2"]
    );
}

#[test]
fn test_strspn_negative_offset() {
    assert_eq!(
        run_prints(
            r#"<?php
echo strspn('foo123456bar', '0123456789', -5), "\n";
"#
        ),
        vec!["2"]
    );
}

#[test]
fn test_strspn_offset_past_end_is_zero() {
    assert_eq!(
        run_prints(
            r#"<?php
echo strspn('abc', 'abc', 10), "\n";
"#
        ),
        vec!["0"]
    );
}

#[test]
fn test_strspn_empty_mask_is_zero() {
    assert_eq!(
        run_prints(
            r#"<?php
echo strspn('abc', ''), "\n";
"#
        ),
        vec!["0"]
    );
}

#[test]
fn test_strspn_negative_offset_beyond_start_is_zero() {
    assert_eq!(
        run_prints(
            r#"<?php
echo strspn('abc123', 'abc', -99), "\n";
"#
        ),
        vec!["0"]
    );
}

#[test]
fn test_strspn_zero_length_subject_returns_zero() {
    assert_eq!(
        run_prints(
            r#"<?php
echo strspn('', 'a', 0), "\n";
echo strspn('', 'a', 0, 10), "\n";
"#
        ),
        vec!["0", "0"]
    );
}

#[test]
fn test_strspn_negative_max_length_truncates() {
    assert_eq!(
        run_prints(
            r#"<?php
echo strspn('12345abc', '123', 0, -2), "\n";
echo strspn('12345abc', '12345', 0, -3), "\n";
"#
        ),
        vec!["3", "5"]
    );
}
