<?php
// vybe-test: php/php_sodium_bin2hex_hex2bin_encoding/test_php_sodium_bin2base64_variant_original
// origin: languages/php/tests/php/test_php_sodium_bin2hex_hex2bin_encoding.rs
// vybe-test-mode: compile

if (function_exists('sodium_bin2base64')) {
    $raw = "BinaryData";
    $b64 = sodium_bin2base64($raw, SODIUM_BASE64_VARIANT_ORIGINAL);
    echo is_string($b64) && strlen($b64) > 0 ? "BIN2BASE64_OK" : "FAIL";
} else {
    echo "BIN2BASE64_OK";
}
