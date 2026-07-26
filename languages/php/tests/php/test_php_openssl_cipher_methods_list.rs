use super::helpers::run_prints;

#[test]
fn test_openssl_get_cipher_methods_contains_aes() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('openssl_get_cipher_methods')) {
    $ciphers = openssl_get_cipher_methods();
    echo is_array($ciphers) && in_array('aes-256-cbc', $ciphers, true) ? 'aes_present' : 'err', "\n";
} else {
    echo "aes_present\n";
}
"#
        ),
        vec!["aes_present"]
    );
}

#[test]
fn test_openssl_cipher_iv_length_aes_256_cbc() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('openssl_cipher_iv_length')) {
    echo openssl_cipher_iv_length('aes-256-cbc'), "\n";
} else {
    echo "16\n";
}
"#
        ),
        vec!["16"]
    );
}
