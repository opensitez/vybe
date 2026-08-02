<?php
// vybe-test: php/php_string_hex_base64_hashing/test_php_hash_raw_output_binary_mode
// origin: languages/php/tests/php/test_php_string_hex_base64_hashing.rs
// vybe-test-mode: compile

$rawHash = hash("sha256", "data", binary: true);
echo strlen($rawHash) === 32 ? "RAW_32_BYTES" : "FAIL";
