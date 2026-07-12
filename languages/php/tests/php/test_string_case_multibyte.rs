use super::helpers::run_prints;

// ── Case conversion functions ─────────────────────────────────

#[test]
fn strtolower_basic() {
    assert_eq!(
        run_prints(r#"<?php echo strtolower('Hello WORLD'); "#),
        vec!["hello world"]
    );
}
#[test]
fn strtoupper_basic() {
    assert_eq!(
        run_prints(r#"<?php echo strtoupper('hello world'); "#),
        vec!["HELLO WORLD"]
    );
}
#[test]
fn ucfirst_basic() {
    assert_eq!(
        run_prints(r#"<?php echo ucfirst('hello world'); "#),
        vec!["Hello world"]
    );
}
#[test]
fn lcfirst_basic() {
    assert_eq!(
        run_prints(r#"<?php echo lcfirst('Hello World'); "#),
        vec!["hello World"]
    );
}
#[test]
fn ucwords_basic() {
    assert_eq!(
        run_prints(r#"<?php echo ucwords('hello world foo'); "#),
        vec!["Hello World Foo"]
    );
}
#[test]
fn ucwords_custom_delimiters() {
    assert_eq!(
        run_prints(r#"<?php echo ucwords('hello-world_foo', '-_'); "#),
        vec!["Hello-World_Foo"]
    );
}

// ── Multibyte case conversion ─────────────────────────────────

#[test]
fn mb_strtolower_unicode() {
    assert_eq!(
        run_prints(r#"<?php echo mb_strtolower('HÉLLO'); "#),
        vec!["héllo"]
    );
}
#[test]
fn mb_strtoupper_unicode() {
    assert_eq!(
        run_prints(r#"<?php echo mb_strtoupper('héllo'); "#),
        vec!["HÉLLO"]
    );
}
#[test]
fn mb_convert_case_title() {
    assert_eq!(
        run_prints(r#"<?php echo mb_convert_case('hello world', MB_CASE_TITLE); "#),
        vec!["Hello World"]
    );
}

// ── String padding and alignment ──────────────────────────────

#[test]
fn mb_str_pad_right() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('mb_str_pad')) {
    echo mb_str_pad('héllo', 8, '.', STR_PAD_RIGHT);
    echo "\n";
} else {
    echo str_pad('hello', 8, '.');
    echo "\n";
}
"#
        ),
        vec!["héllo..."]
    );
}

// ── mb_strlen vs strlen ───────────────────────────────────────

#[test]
fn mb_strlen_vs_strlen() {
    assert_eq!(
        run_prints(
            r#"<?php
$s = 'héllo';
echo strlen($s) . ':' . mb_strlen($s);
echo "\n";
"#
        ),
        vec!["6:5"]
    );
}

// ── mb_substr ────────────────────────────────────────────────

#[test]
fn mb_substr_from_unicode() {
    assert_eq!(
        run_prints(r#"<?php echo mb_substr('hello wörld', 0, 5); "#),
        vec!["hello"]
    );
}
#[test]
fn mb_substr_negative_offset() {
    assert_eq!(
        run_prints(r#"<?php echo mb_substr('héllo', -3); "#),
        vec!["llo"]
    );
}
#[test]
fn mb_substr_negative_length() {
    assert_eq!(
        run_prints(r#"<?php echo mb_substr('héllo', 0, -2); "#),
        vec!["hél"]
    );
}

// ── mb_strpos / mb_strrpos ────────────────────────────────────

#[test]
fn mb_strpos_finds_char() {
    assert_eq!(
        run_prints(r#"<?php echo mb_strpos('héllo', 'l'); "#),
        vec!["2"]
    );
}
#[test]
fn mb_strrpos_last_occurrence() {
    assert_eq!(
        run_prints(r#"<?php echo mb_strrpos('héllo', 'l'); "#),
        vec!["3"]
    );
}
#[test]
fn mb_strpos_not_found() {
    assert_eq!(
        run_prints(r#"<?php var_export(mb_strpos('héllo', 'x')); "#),
        vec!["false"]
    );
}

// ── mb_detect_encoding ────────────────────────────────────────

#[test]
fn mb_detect_encoding_utf8() {
    assert_eq!(
        run_prints(
            r#"<?php
$s = 'hello';
echo mb_check_encoding($s, 'UTF-8') ? 'utf8' : 'not';
echo "\n";
"#
        ),
        vec!["utf8"]
    );
}

// ── mb_str_split ─────────────────────────────────────────────

#[test]
fn mb_str_split_default() {
    assert_eq!(
        run_prints(r#"<?php echo implode(',', mb_str_split('abc')); "#),
        vec!["a,b,c"]
    );
}
#[test]
fn mb_str_split_with_length() {
    assert_eq!(
        run_prints(r#"<?php echo implode('|', mb_str_split('abcdef', 2)); "#),
        vec!["ab|cd|ef"]
    );
}

// ── mb_substr_count ───────────────────────────────────────────

#[test]
fn mb_substr_count_unicode() {
    assert_eq!(
        run_prints(r#"<?php echo mb_substr_count('héllo héllo', 'é'); "#),
        vec!["2"]
    );
}

// ── String comparison case-insensitive ───────────────────────

#[test]
fn strcasecmp_equal() {
    assert_eq!(
        run_prints(r#"<?php echo strcasecmp('Hello', 'hello') === 0 ? 'eq' : 'neq'; "#),
        vec!["eq"]
    );
}
#[test]
fn strncasecmp_prefix() {
    assert_eq!(
        run_prints(r#"<?php echo strncasecmp('Hello World', 'HELLO', 5) === 0 ? 'match' : 'no'; "#),
        vec!["match"]
    );
}
