use super::helpers::run_prints;

#[test]
fn test_mb_strwidth_full_width_characters() {
    assert_eq!(
        run_prints(
            r#"<?php
echo mb_strwidth('hello') . ',' . mb_strwidth('日本語'), "\n";
"#
        ),
        vec!["5,6"]
    );
}

#[test]
fn test_mb_strimwidth_custom_trim_marker() {
    assert_eq!(
        run_prints(
            r#"<?php
echo mb_strimwidth('Hello World', 0, 8, '...'), "\n";
"#
        ),
        vec!["Hello..."]
    );
}

#[test]
fn test_mb_strimwidth_multibyte_trim() {
    assert_eq!(
        run_prints(
            r#"<?php
echo mb_strimwidth('東京都千代田区', 0, 8, '..'), "\n";
"#
        ),
        vec!["東京.."]
    );
}

#[test]
fn test_mb_strimwidth_zero_width_returns_empty() {
    assert_eq!(
        run_prints(
            r#"<?php
echo mb_strimwidth('abcdef', 0, 0, '...');
"#
        ),
        vec![""]
    );
}

#[test]
fn test_mb_strimwidth_with_offset_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo mb_strimwidth('abcdef', 2, 4, '..'), "\n";
echo mb_strimwidth('日本語テスト', 1, 4, '..');
"#
        ),
        vec!["cd..|本.."]
    );
}
