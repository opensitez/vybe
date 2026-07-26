use super::helpers::run_prints;

#[test]
fn test_mb_ord_ascii_and_multibyte() {
    assert_eq!(
        run_prints(
            r#"<?php
echo mb_ord('A') . ',' . mb_ord('€'), "\n";
"#
        ),
        vec!["65,8364"]
    );
}

#[test]
fn test_mb_chr_code_point_to_string() {
    assert_eq!(
        run_prints(
            r#"<?php
echo mb_chr(65) . ',' . mb_chr(8364), "\n";
"#
        ),
        vec!["A,€"]
    );
}

#[test]
fn test_mb_ord_mb_chr_roundtrip() {
    assert_eq!(
        run_prints(
            r#"<?php
$char = '𝄢'; // Musical symbol F clef
$code = mb_ord($char);
$restored = mb_chr($code);
echo ($char === $restored) ? 'roundtrip_ok' : 'err', "\n";
"#
        ),
        vec!["roundtrip_ok"]
    );
}
