<?php
// vybe-test: php/php_password_hashing_cryptography/test_php_hash_hkdf_key_derivation
// origin: languages/php/tests/php/test_php_password_hashing_cryptography.rs
// vybe-test-mode: compile

$ikm = "input_key_material";
$derived = hash_hkdf("sha256", $ikm, 32, "info_label");
echo strlen($derived) === 32 ? "HKDF_32_BYTES" : "FAIL";
