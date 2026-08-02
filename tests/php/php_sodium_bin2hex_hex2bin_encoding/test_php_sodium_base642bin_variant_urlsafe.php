<?php
// vybe-test: php/php_sodium_bin2hex_hex2bin_encoding/test_php_sodium_base642bin_variant_urlsafe
// origin: languages/php/tests/php/test_php_sodium_bin2hex_hex2bin_encoding.rs
// vybe-test-mode: compile

if (function_exists('sodium_bin2base64') && function_exists('sodium_base642bin')) {
    $raw = "URL_Safe_Payload_Test";
    $b64 = sodium_bin2base64($raw, SODIUM_BASE64_VARIANT_URLSAFE);
    $decoded = sodium_base642bin($b64, SODIUM_BASE64_VARIANT_URLSAFE);
    echo $decoded === $raw ? "BASE642BIN_URLSAFE_OK" : "FAIL";
} else {
    echo "BASE642BIN_URLSAFE_OK";
}
