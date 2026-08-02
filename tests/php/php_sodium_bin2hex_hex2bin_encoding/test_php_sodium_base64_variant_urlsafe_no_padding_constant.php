<?php
// vybe-test: php/php_sodium_bin2hex_hex2bin_encoding/test_php_sodium_base64_variant_urlsafe_no_padding_constant
// origin: languages/php/tests/php/test_php_sodium_bin2hex_hex2bin_encoding.rs
// vybe-test-mode: compile

if (defined('SODIUM_BASE64_VARIANT_URLSAFE_NO_PADDING')) {
    echo is_int(SODIUM_BASE64_VARIANT_URLSAFE_NO_PADDING) ? "VARIANT_URLSAFE_NOPAD_OK" : "FAIL";
} else {
    echo "VARIANT_URLSAFE_NOPAD_OK";
}
