use super::helpers::run_prints;

// ── strtr ─────────────────────────────────────────────────────

#[test]
fn strtr_single_char_map() {
    assert_eq!(
        run_prints(r#"<?php echo strtr('hello world', 'aeiou', '*****'); "#),
        vec!["h*ll* w*rld"]
    );
}
#[test]
fn strtr_array_map() {
    assert_eq!(
        run_prints(r#"<?php echo strtr('php is cool', ['php'=>'Vybe','cool'=>'awesome']); "#),
        vec!["Vybe is awesome"]
    );
}
#[test]
fn strtr_longer_match_wins() {
    assert_eq!(
        run_prints(r#"<?php echo strtr('aa', ['a'=>'b','aa'=>'c']); "#),
        vec!["c"]
    );
}

// ── wordwrap ──────────────────────────────────────────────────

#[test]
fn wordwrap_basic() {
    assert_eq!(
        run_prints(r#"<?php echo wordwrap('The quick brown fox', 10, "\n"); "#),
        vec!["The quick\nbrown fox"]
    );
}
#[test]
fn wordwrap_cut_long_words() {
    assert_eq!(
        run_prints(r#"<?php echo wordwrap('superlongword', 5, '-', true); "#),
        vec!["super-longw-ord"]
    );
}

// ── chunk_split ───────────────────────────────────────────────

#[test]
fn chunk_split_hex_display() {
    assert_eq!(
        run_prints(r#"<?php echo rtrim(chunk_split('AABBCCDD', 2, ':'), ':'); "#),
        vec!["AA:BB:CC:DD"]
    );
}
#[test]
fn chunk_split_base64_style() {
    assert_eq!(
        run_prints(r#"<?php echo chunk_split('abcdefghij', 4, '-'); "#),
        vec!["abcd-efgh-ij-"]
    );
}

// ── str_pad ───────────────────────────────────────────────────

#[test]
fn str_pad_right_default() {
    assert_eq!(
        run_prints(r#"<?php echo str_pad('42', 6); "#),
        vec!["42    "]
    );
}
#[test]
fn str_pad_left() {
    assert_eq!(
        run_prints(r#"<?php echo str_pad('42', 6, '0', STR_PAD_LEFT); "#),
        vec!["000042"]
    );
}
#[test]
fn str_pad_both() {
    assert_eq!(
        run_prints(r#"<?php echo str_pad('hi', 8, '-', STR_PAD_BOTH); "#),
        vec!["---hi---"]
    );
}
#[test]
fn str_pad_custom_char() {
    assert_eq!(
        run_prints(r#"<?php echo str_pad('x', 5, '*'); "#),
        vec!["x****"]
    );
}
#[test]
fn str_pad_shorter_than_input_unchanged() {
    assert_eq!(
        run_prints(r#"<?php echo str_pad('hello', 3); "#),
        vec!["hello"]
    );
}

// ── nl2br ─────────────────────────────────────────────────────

#[test]
fn nl2br_inserts_br_before_newline() {
    assert_eq!(
        run_prints(r#"<?php echo nl2br("line1\nline2"); "#),
        vec!["line1<br />\nline2"]
    );
}
#[test]
fn nl2br_xhtml_false_gives_html4() {
    assert_eq!(
        run_prints(r#"<?php echo nl2br("a\nb", false); "#),
        vec!["a<br>\nb"]
    );
}

// ── str_repeat ────────────────────────────────────────────────

#[test]
fn str_repeat_basic() {
    assert_eq!(
        run_prints(r#"<?php echo str_repeat('ab', 3); "#),
        vec!["ababab"]
    );
}
#[test]
fn str_repeat_zero_times() {
    assert_eq!(run_prints(r#"<?php echo str_repeat('x', 0); "#), vec![""]);
}

// ── number_format / money_format patterns ────────────────────

#[test]
fn number_format_thousands_comma() {
    assert_eq!(
        run_prints(r#"<?php echo number_format(1234567.891, 2); "#),
        vec!["1,234,567.89"]
    );
}
#[test]
fn number_format_custom_separators() {
    assert_eq!(
        run_prints(r#"<?php echo number_format(1234567.5, 2, ',', '.'); "#),
        vec!["1.234.567,50"]
    );
}

// ── printf / sprintf ─────────────────────────────────────────

#[test]
fn printf_returns_length() {
    assert_eq!(
        run_prints(r#"<?php $len = printf('%s', 'hello'); echo ' ' . $len; "#),
        vec!["hello 5"]
    );
}

// ── str_contains / str_starts_with / str_ends_with ───────────

#[test]
fn str_contains_true() {
    assert_eq!(
        run_prints(r#"<?php echo str_contains('Hello World', 'World') ? 'yes' : 'no'; "#),
        vec!["yes"]
    );
}
#[test]
fn str_starts_with_true() {
    assert_eq!(
        run_prints(r#"<?php echo str_starts_with('PHP 8.3', 'PHP') ? 'yes' : 'no'; "#),
        vec!["yes"]
    );
}
#[test]
fn str_ends_with_true() {
    assert_eq!(
        run_prints(r#"<?php echo str_ends_with('hello.php', '.php') ? 'yes' : 'no'; "#),
        vec!["yes"]
    );
}
#[test]
fn str_contains_empty_needle() {
    assert_eq!(
        run_prints(r#"<?php echo str_contains('anything', '') ? 'yes' : 'no'; "#),
        vec!["yes"]
    );
}

// ── substr_count / substr_replace ────────────────────────────

#[test]
fn substr_count_basic() {
    assert_eq!(
        run_prints(r#"<?php echo substr_count('hello world hello', 'hello'); "#),
        vec!["2"]
    );
}
#[test]
fn substr_replace_basic() {
    assert_eq!(
        run_prints(r#"<?php echo substr_replace('Hello World', 'PHP', 6, 5); "#),
        vec!["Hello PHP"]
    );
}
#[test]
fn substr_replace_negative_offset() {
    assert_eq!(
        run_prints(r#"<?php echo substr_replace('Hello World', '!', -5, 5); "#),
        vec!["Hello !"]
    );
}
