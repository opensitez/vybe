<?php
// vybe-test: php/php_string_hex_base64_hashing/test_php_hash_init_update_final_incremental
// origin: languages/php/tests/php/test_php_string_hex_base64_hashing.rs
// vybe-test-mode: compile

$ctx = hash_init("sha256");
hash_update($ctx, "Part 1 ");
hash_update($ctx, "Part 2");
$digest = hash_final($ctx);

echo strlen($digest) === 64 ? "INCREMENTAL_HASH_OK" : "FAIL";
