<?php
// vybe-test: php/hash_crypto/crc32_hash
// origin: languages/php/tests/php/test_hash_crypto.rs
// vybe-test-mode: compile

$checksum = crc32('hello world');
echo is_int($checksum) ? 'int' : 'not int';
echo crc32('hello world') === crc32('hello world') ? ':deterministic' : ':varies';
