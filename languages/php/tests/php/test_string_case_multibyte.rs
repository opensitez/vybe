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
fn mb_strtolower_empty_string() {
    assert_eq!(
        run_prints(r#"<?php echo var_export(mb_strtolower(''), true); "#),
        vec!["''"]
    );
}

#[test]
fn mb_strtoupper_with_digits_and_spaces() {
    assert_eq!(
        run_prints(r#"<?php echo mb_strtoupper('x y 1a b'); "#),
        vec!["X Y 1A B"]
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

#[test]
fn mb_convert_case_lower_and_upper_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo mb_convert_case('AbC', MB_CASE_LOWER);
echo '|';
echo mb_convert_case('AbC', MB_CASE_UPPER);
"#
        ),
        vec!["abc|ABC"]
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

#[test]
fn mb_str_pad_custom_padding() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('mb_str_pad')) {
    echo mb_str_pad('é', 3, '0', STR_PAD_LEFT);
} else {
    echo str_pad('é', 3, '0', STR_PAD_LEFT);
}
echo '|';
if (function_exists('mb_str_pad')) {
    echo mb_str_pad('é', 2, '*', STR_PAD_BOTH);
} else {
    echo str_pad('é', 2, '*', STR_PAD_BOTH);
}
"#
        ),
        vec!["00é|*é"]
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

#[test]
fn mb_strlen_empty_and_numeric() {
    assert_eq!(
        run_prints(r#"<?php echo mb_strlen('') . '|' . mb_strlen('123'); "#),
        vec!["0|3"]
    );
}

#[test]
fn mb_strlen_zero_width_subject() {
    assert_eq!(
        run_prints(
            r#"<?php
$s = "\0";
echo mb_strlen($s);
"#
        ),
        vec!["1"]
    );
}

#[test]
fn mb_strlen_of_multiline_unicode() {
    assert_eq!(
        run_prints(
            r#"<?php
$s = "hé\n";
echo strlen($s) . "|" . mb_strlen($s);
"#
        ),
        vec!["4|3"]
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

#[test]
fn mb_str_split_longer_length_returns_single_chunk() {
    assert_eq!(
        run_prints(
            r#"<?php echo count(mb_str_split('hey', 10)); echo '|'; echo mb_str_split('hey', 10)[0]; "#
        ),
        vec!["1|hey"]
    );
}

#[test]
fn mb_check_encoding_false_branch() {
    assert_eq!(
        run_prints(
            r#"<?php
$invalid = chr(255);
echo mb_check_encoding($invalid, 'UTF-8') ? 'ok' : 'bad';
echo '|';
echo mb_check_encoding('test', 'ISO-8859-1') ? 'lat' : 'no';
"#
        ),
        vec!["bad|lat"]
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
