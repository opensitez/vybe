<?php
// vybe-test: php/php_string_hex_base64_hashing/test_php_hash_hmac_file_checksum
// origin: languages/php/tests/php/test_php_string_hex_base64_hashing.rs
// vybe-test-mode: compile

$tmp = tempnam(sys_get_temp_dir(), "hash_file_");
file_put_contents($tmp, "file content for hmac");
$hmac = hash_hmac_file("sha256", $tmp, "secret_key");
unlink($tmp);

echo strlen($hmac) === 64 ? "HMAC_FILE_OK" : "FAIL";
