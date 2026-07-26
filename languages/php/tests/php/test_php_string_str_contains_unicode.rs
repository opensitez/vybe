use super::helpers::run_prints;

#[test]
fn test_str_contains_utf8_multibyte() {
    assert_eq!(
        run_prints(
            r#"<?php
echo str_contains('café', 'fé') ? 'found' : 'not_found', "\n";
"#
        ),
        vec!["found"]
    );
}

#[test]
fn test_str_contains_emoji() {
    assert_eq!(
        run_prints(
            r#"<?php
echo str_contains('hello 😃 world', '😃') ? 'emoji_found' : 'missing', "\n";
"#
        ),
        vec!["emoji_found"]
    );
}

#[test]
fn test_str_contains_empty_needle() {
    assert_eq!(
        run_prints(
            r#"<?php
echo str_contains('anything', '') ? 'empty_needle_matches' : 'no', "\n";
"#
        ),
        vec!["empty_needle_matches"]
    );
}

#[test]
fn test_str_contains_case_sensitive() {
    assert_eq!(
        run_prints(
            r#"<?php
echo str_contains('PHP 8', 'php') ? 'match' : 'no_match', "\n";
"#
        ),
        vec!["no_match"]
    );
}

#[test]
fn test_str_starts_with_unicode_char() {
    assert_eq!(
        run_prints(
            r#"<?php
echo str_starts_with('café noir', 'café') ? 'starts' : 'no';
echo "\n";
echo str_starts_with('café noir', 'CAFÉ') ? 'starts-upper' : 'no-upper';
"#
        ),
        vec!["starts|no-upper"]
    );
}

#[test]
fn test_str_ends_with_unicode_char() {
    assert_eq!(
        run_prints(
            r#"<?php
echo str_ends_with('smile 😃', '😃') ? 'ends' : 'no';
echo "\n";
echo str_ends_with('smile 😃', '') ? 'empty' : 'no-empty';
"#
        ),
        vec!["ends|empty"]
    );
}

#[test]
fn test_str_starts_with_empty_needle_false_subject() {
    assert_eq!(
        run_prints(
            r#"<?php
echo str_starts_with('', 'x') ? 'yes' : 'no';
echo "\n";
echo str_starts_with('', '') ? 'yes2' : 'no2';
echo "\n";
echo str_ends_with('', '') ? 'yes3' : 'no3';
"#
        ),
        vec!["no|yes2|yes3"]
    );
}
