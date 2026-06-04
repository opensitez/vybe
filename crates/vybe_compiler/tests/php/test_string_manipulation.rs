use super::helpers::run_prints;

// ── String search functions ───────────────────────────────────

#[test]
fn strpos_found() {
    assert_eq!(
        run_prints(r#"<?php echo strpos('Hello World', 'World'); "#),
        vec!["6"]
    );
}
#[test]
fn strpos_not_found_returns_false() {
    assert_eq!(
        run_prints(r#"<?php var_export(strpos('Hello', 'xyz')); "#),
        vec!["false"]
    );
}
#[test]
fn strrpos_last_occurrence() {
    assert_eq!(
        run_prints(r#"<?php echo strrpos('hello world hello', 'hello'); "#),
        vec!["12"]
    );
}
#[test]
fn strripos_case_insensitive_last() {
    assert_eq!(
        run_prints(r#"<?php echo strripos('Hello World Hello', 'hello'); "#),
        vec!["12"]
    );
}
#[test]
fn substr_count_with_offset() {
    assert_eq!(
        run_prints(r#"<?php echo substr_count('aaaa', 'aa'); "#),
        vec!["2"]
    );
}

// ── String modification ───────────────────────────────────────

#[test]
fn ltrim_removes_left() {
    assert_eq!(
        run_prints(r#"<?php echo ltrim('   hello   '); "#),
        vec!["hello   "]
    );
}
#[test]
fn rtrim_removes_right() {
    assert_eq!(
        run_prints(r#"<?php echo rtrim('   hello   '); "#),
        vec!["   hello"]
    );
}
#[test]
fn trim_custom_chars() {
    assert_eq!(
        run_prints(r#"<?php echo trim('***hello***', '*'); "#),
        vec!["hello"]
    );
}
#[test]
fn str_rot13() {
    assert_eq!(
        run_prints(r#"<?php echo str_rot13(str_rot13('Hello World')); "#),
        vec!["Hello World"]
    );
}
#[test]
fn strrev_basic() {
    assert_eq!(run_prints(r#"<?php echo strrev('Hello'); "#), vec!["olleH"]);
}

// ── String conversion ─────────────────────────────────────────

#[test]
fn bin2hex_hex2bin_roundtrip() {
    assert_eq!(
        run_prints(r#"<?php echo hex2bin(bin2hex('Hello')); "#),
        vec!["Hello"]
    );
}
#[test]
fn ord_chr_roundtrip() {
    assert_eq!(run_prints(r#"<?php echo chr(ord('A')); "#), vec!["A"]);
}
#[test]
fn crc32_consistent() {
    assert_eq!(
        run_prints(r#"<?php echo crc32('hello') === crc32('hello') ? 'same' : 'diff'; "#),
        vec!["same"]
    );
}

// ── String padding and alignment ──────────────────────────────

#[test]
fn str_pad_right_with_string() {
    assert_eq!(
        run_prints(r#"<?php echo str_pad('42', 10, '0', STR_PAD_LEFT); "#),
        vec!["0000000042"]
    );
}
#[test]
fn str_pad_repeating_pad() {
    assert_eq!(
        run_prints(r#"<?php echo str_pad('x', 9, 'abc'); "#),
        vec!["xabcabcab"]
    );
}

// ── Tokenizing and parsing ────────────────────────────────────

#[test]
fn strtok_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$s = 'hello world foo';
$token = strtok($s, ' ');
$tokens = [];
while ($token !== false) { $tokens[] = $token; $token = strtok(' '); }
echo implode(',', $tokens);
echo "\n";
"#
        ),
        vec!["hello,world,foo"]
    );
}
#[test]
fn sscanf_parse_date() {
    assert_eq!(
        run_prints(
            r#"<?php
[$y,$m,$d] = sscanf('2024-07-15', '%d-%d-%d');
echo "$d/$m/$y";
echo "\n";
"#
        ),
        vec!["15/7/2024"]
    );
}

// ── String matching ───────────────────────────────────────────

#[test]
fn fnmatch_wildcard() {
    assert_eq!(
        run_prints(r#"<?php echo fnmatch('*.php', 'index.php') ? 'yes' : 'no'; "#),
        vec!["yes"]
    );
}
#[test]
fn fnmatch_no_match() {
    assert_eq!(
        run_prints(r#"<?php echo fnmatch('*.php', 'index.html') ? 'yes' : 'no'; "#),
        vec!["no"]
    );
}

// ── String info ───────────────────────────────────────────────

#[test]
fn str_word_count_mode0() {
    assert_eq!(
        run_prints(r#"<?php echo str_word_count('Hello World PHP'); "#),
        vec!["3"]
    );
}
#[test]
fn str_word_count_mode1() {
    assert_eq!(
        run_prints(r#"<?php echo implode(',', str_word_count('one two three', 1)); "#),
        vec!["one,two,three"]
    );
}
#[test]
fn str_word_count_mode2() {
    assert_eq!(
        run_prints(r#"<?php $m = str_word_count('hello world', 2); echo $m[0] . ',' . $m[6]; "#),
        vec!["hello,world"]
    );
}
#[test]
fn similar_text_percentage() {
    assert_eq!(
        run_prints(r#"<?php similar_text('World', 'Word', $p); echo round($p); "#),
        vec!["89"]
    );
}
#[test]
fn levenshtein_distance() {
    assert_eq!(
        run_prints(r#"<?php echo levenshtein('kitten', 'sitting'); "#),
        vec!["3"]
    );
}

// ── Multibyte string functions ────────────────────────────────

#[test]
fn mb_strlen_unicode() {
    assert_eq!(run_prints(r#"<?php echo mb_strlen('héllo'); "#), vec!["5"]);
}
#[test]
fn mb_strtoupper_unicode() {
    assert_eq!(
        run_prints(r#"<?php echo mb_strtoupper('héllo'); "#),
        vec!["HÉLLO"]
    );
}
#[test]
fn mb_substr_unicode() {
    assert_eq!(
        run_prints(r#"<?php echo mb_substr('héllo wörld', 6, 5); "#),
        vec!["wörld"]
    );
}
#[test]
fn mb_strpos_unicode() {
    assert_eq!(
        run_prints(r#"<?php echo mb_strpos('héllo wörld', 'wörld'); "#),
        vec!["6"]
    );
}
