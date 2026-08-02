<?php
// vybe-test: php/php_string_hex_base64_hashing/test_php_hash_file_digest
// origin: languages/php/tests/php/test_php_string_hex_base64_hashing.rs
// vybe-test-mode: compile

$tmp = tempnam(sys_get_temp_dir(), "hash_digest_");
file_put_contents($tmp, "content");
$digest = hash_file("md5", $tmp);
unlink($tmp);

echo strlen($digest) === 32 ? "HASH_FILE_OK" : "FAIL";
