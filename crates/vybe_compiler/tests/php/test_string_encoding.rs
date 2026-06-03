use super::helpers::run_prints;

// ── URL encoding ─────────────────────────────────────────────

#[test]
fn urlencode_basic() {
    assert_eq!(
        run_prints(r#"<?php echo urlencode('hello world'); "#),
        vec!["hello+world"]
    );
}
#[test]
fn urldecode_basic() {
    assert_eq!(
        run_prints(r#"<?php echo urldecode('hello+world'); "#),
        vec!["hello world"]
    );
}
#[test]
fn rawurlencode_uses_percent() {
    assert_eq!(
        run_prints(r#"<?php echo rawurlencode('hello world'); "#),
        vec!["hello%20world"]
    );
}
#[test]
fn rawurldecode_percent_space() {
    assert_eq!(
        run_prints(r#"<?php echo rawurldecode('hello%20world'); "#),
        vec!["hello world"]
    );
}
#[test]
fn urlencode_special_chars() {
    assert_eq!(
        run_prints(r#"<?php echo urlencode('a=1&b=2'); "#),
        vec!["a%3D1%26b%3D2"]
    );
}
#[test]
fn urlencode_roundtrip() {
    assert_eq!(
        run_prints(r#"<?php $s = 'foo bar+baz'; echo urldecode(urlencode($s)); "#),
        vec!["foo bar+baz"]
    );
}

// ── HTML encoding ─────────────────────────────────────────────

#[test]
fn htmlspecialchars_escapes_ampersand() {
    assert_eq!(
        run_prints(r#"<?php echo htmlspecialchars('Tom & Jerry'); "#),
        vec!["Tom &amp; Jerry"]
    );
}
#[test]
fn htmlspecialchars_escapes_tags() {
    assert_eq!(
        run_prints(r#"<?php echo htmlspecialchars('<script>alert(1)</script>'); "#),
        vec!["&lt;script&gt;alert(1)&lt;/script&gt;"]
    );
}
#[test]
fn htmlspecialchars_decode() {
    assert_eq!(
        run_prints(r#"<?php echo htmlspecialchars_decode('&lt;b&gt;hello&lt;/b&gt;'); "#),
        vec!["<b>hello</b>"]
    );
}
#[test]
fn html_entity_decode() {
    assert_eq!(
        run_prints(r#"<?php echo html_entity_decode('&amp;&lt;&gt;&quot;'); "#),
        vec!["<>&\""]
    );
}
#[test]
fn htmlentities_non_ascii() {
    assert_eq!(
        run_prints(r#"<?php echo htmlentities('café', ENT_QUOTES, 'UTF-8'); "#),
        vec!["caf&eacute;"]
    );
}

// ── Base64 encoding ───────────────────────────────────────────

#[test]
fn base64_encode_decode_roundtrip() {
    assert_eq!(
        run_prints(r#"<?php $s = 'Hello, World!'; echo base64_decode(base64_encode($s)); "#),
        vec!["Hello, World!"]
    );
}
#[test]
fn base64_encode_standard() {
    assert_eq!(
        run_prints(r#"<?php echo base64_encode('Man'); "#),
        vec!["TWFu"]
    );
}
#[test]
fn base64_decode_invalid_returns_false() {
    assert_eq!(
        run_prints(r#"<?php $r = base64_decode('not!base64!!!', true); var_export($r); "#),
        vec!["false"]
    );
}

// ── Character conversion ──────────────────────────────────────

#[test]
fn ord_and_chr() {
    assert_eq!(
        run_prints(r#"<?php echo ord('A') . ',' . chr(97); "#),
        vec!["65,a"]
    );
}
#[test]
fn chr_sequence() {
    assert_eq!(
        run_prints(r#"<?php echo implode('', array_map('chr', range(65, 69))); "#),
        vec!["ABCDE"]
    );
}

// ── Numeric string detection ──────────────────────────────────

#[test]
fn ctype_digit() {
    assert_eq!(
        run_prints(r#"<?php echo ctype_digit('12345') ? 'yes' : 'no'; "#),
        vec!["yes"]
    );
}
#[test]
fn ctype_digit_with_sign() {
    assert_eq!(
        run_prints(r#"<?php echo ctype_digit('-123') ? 'yes' : 'no'; "#),
        vec!["no"]
    );
}
#[test]
fn ctype_alpha() {
    assert_eq!(
        run_prints(r#"<?php echo ctype_alpha('Hello') ? 'yes' : 'no'; "#),
        vec!["yes"]
    );
}
#[test]
fn ctype_alnum() {
    assert_eq!(
        run_prints(r#"<?php echo ctype_alnum('Hello123') ? 'yes' : 'no'; "#),
        vec!["yes"]
    );
}
#[test]
fn ctype_space() {
    assert_eq!(
        run_prints(r#"<?php echo ctype_space("   \t\n") ? 'yes' : 'no'; "#),
        vec!["yes"]
    );
}
#[test]
fn ctype_upper_lower() {
    assert_eq!(
        run_prints(
            r#"<?php echo ctype_upper('HELLO') ? '1' : '0'; echo ctype_lower('world') ? '1' : '0'; "#
        ),
        vec!["11"]
    );
}

// ── Hash functions ────────────────────────────────────────────

#[test]
fn md5_consistent() {
    assert_eq!(
        run_prints(r#"<?php echo md5('hello'); "#),
        vec!["5d41402abc4b2a76b9719d911017c592"]
    );
}
#[test]
fn sha1_consistent() {
    assert_eq!(
        run_prints(r#"<?php echo sha1('hello'); "#),
        vec!["aaf4c61ddcc5e8a2dabede0f3b482cd9aea9434d"]
    );
}
#[test]
fn hash_sha256() {
    assert_eq!(
        run_prints(r#"<?php echo strlen(hash('sha256', 'hello')); "#),
        vec!["64"]
    );
}
#[test]
fn hash_hmac() {
    assert_eq!(
        run_prints(r#"<?php $hmac = hash_hmac('sha256', 'message', 'key'); echo strlen($hmac); "#),
        vec!["64"]
    );
}
