use super::helpers::run_prints;

#[test]
fn test_strspn_initial_segment_length() {
    assert_eq!(
        run_prints(
            r#"<?php
echo strspn('123456abcdef', '0123456789'), "\n";
"#
        ),
        vec!["6"]
    );
}

#[test]
fn test_strspn_with_start_offset() {
    assert_eq!(
        run_prints(
            r#"<?php
echo strspn('abc123def', '0123456789', 3), "\n";
"#
        ),
        vec!["3"]
    );
}

#[test]
fn test_strcspn_complement_mask_length() {
    assert_eq!(
        run_prints(
            r#"<?php
echo strcspn('hello world', 'w'), "\n";
"#
        ),
        vec!["6"]
    );
}

#[test]
fn test_strcspn_with_offset_and_length() {
    assert_eq!(
        run_prints(
            r#"<?php
echo strcspn('foo bar baz', 'z', 4, 5), "\n";
"#
        ),
        vec!["5"]
    );
}

#[test]
fn test_strspn_empty_mask_for_non_empty_subject() {
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
fn test_strcspn_no_matching_mask_is_full_length() {
    assert_eq!(
        run_prints(
            r#"<?php
echo strcspn('abc', 'x'), "\n";
"#
        ),
        vec!["3"]
    );
}

#[test]
fn test_strcspn_negative_offset() {
    assert_eq!(
        run_prints(
            r#"<?php
echo strcspn('a1b2c3', '123', -4), "\n";
"#
        ),
        vec!["0"]
    );
}
