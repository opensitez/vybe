<?php
// vybe-test: php/php_openssl_cipher_iv_length_lookup/test_php_openssl_get_md_methods_list
// origin: languages/php/tests/php/test_php_openssl_cipher_iv_length_lookup.rs
// vybe-test-mode: compile

if (function_exists('openssl_get_md_methods')) {
    $mds = openssl_get_md_methods();
    echo in_array("sha256", $mds) || in_array("SHA256", $mds) ? "SHA256_AVAILABLE" : "FAIL";
} else {
    echo "SHA256_AVAILABLE";
}
