use super::helpers::run_prints;

#[test]
fn test_hash_hkdf_length_and_hex() {
    assert_eq!(
        run_prints(
            r#"<?php
$derived = hash_hkdf('sha256', 'secret_ikm', 32, 'app_info', 'salt123');
echo strlen($derived) . ':' . bin2hex(substr($derived, 0, 4)), "\n";
"#
        ),
        vec!["32:a816bd4d"]
    );
}

#[test]
fn test_hash_pbkdf2_raw_output_true() {
    assert_eq!(
        run_prints(
            r#"<?php
$raw = hash_pbkdf2('sha256', 'password', 'salt', 1000, 16, true);
echo strlen($raw) . ':' . gettype($raw), "\n";
"#
        ),
        vec!["16:string"]
    );
}

#[test]
fn test_hash_pbkdf2_hex_output() {
    assert_eq!(
        run_prints(
            r#"<?php
$hex = hash_pbkdf2('sha256', 'password', 'salt', 1000, 20, false);
echo strlen($hex), "\n";
"#
        ),
        vec!["20"]
    );
}
