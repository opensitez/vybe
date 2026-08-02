<?php
// vybe-test: php/php_string_hex_base64_hashing/test_php_base64_decode_strict_mode
// origin: languages/php/tests/php/test_php_string_hex_base64_hashing.rs
// vybe-test-mode: compile

$invalidBase64 = "===invalid===";
$res = base64_decode($invalidBase64, strict: true);
echo $res === false ? "STRICT_DECODE_FAILED" : "DECODED";
