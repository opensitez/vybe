<?php
// vybe-test: php/php_string_hex_base64_hashing/test_php_hash_copy_context
// origin: languages/php/tests/php/test_php_string_hex_base64_hashing.rs
// vybe-test-mode: compile

$ctx1 = hash_init("md5");
hash_update($ctx1, "hello");
$ctx2 = hash_copy($ctx1);
hash_update($ctx1, " world");
hash_update($ctx2, " php");

echo hash_final($ctx1) !== hash_final($ctx2) ? "DIFFERENT_HASHES" : "SAME";
