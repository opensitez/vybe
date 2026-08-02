<?php
// vybe-test: php/hash_crypto/hash_md5_via_hash
// origin: languages/php/tests/php/test_hash_crypto.rs
// vybe-test-mode: compile

$h = hash('md5', 'hello');
echo strlen($h) === 32 ? 'ok' : 'fail';
echo $h === md5('hello') ? ':matches' : ':differs';
