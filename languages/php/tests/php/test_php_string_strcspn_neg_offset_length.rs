use super::helpers::run_prints;

#[test]
fn test_strcspn_negative_offset() {
    assert_eq!(
        run_prints(
            r#"<?php
echo strcspn('hello world', 'o', -6), "\n";
"#
        ),
        vec!["1"]
    );
}

#[test]
fn test_strcspn_negative_length() {
    assert_eq!(
        run_prints(
            r#"<?php
echo strcspn('foo bar baz', 'z', 0, -2), "\n";
"#
        ),
        vec!["9"]
    );
}

#[test]
fn test_strcspn_empty_mask_scans_full_length() {
    assert_eq!(
        run_prints(
            r#"<?php
echo strcspn('abc', ''), "\n";
"#
        ),
        vec!["3"]
    );
}

#[test]
fn test_strcspn_offset_beyond_end_is_zero() {
    assert_eq!(
        run_prints(
            r#"<?php
echo strcspn('abcdef', 'a', 10), "\n";
"#
        ),
        vec!["0"]
    );
}
