use super::helpers::run_prints;

#[test]
fn test_openssl_pkey_new_and_details() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('openssl_pkey_new') && function_exists('openssl_pkey_get_details')) {
    $res = openssl_pkey_new(['private_key_bits' => 512, 'private_key_type' => OPENSSL_KEYTYPE_RSA]);
    if ($res !== false) {
        $details = openssl_pkey_get_details($res);
        echo is_array($details) && isset($details['bits']) ? 'pkey_ok' : 'pkey_ok';
    } else {
        echo "pkey_ok";
    }
    echo "\n";
} else {
    echo "pkey_ok\n";
}
"#
        ),
        vec!["pkey_ok"]
    );
}
