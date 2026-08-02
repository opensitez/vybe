<?php
// vybe-test: php/hash_crypto/hash_sha1_via_hash
// origin: languages/php/tests/php/test_hash_crypto.rs
// vybe-test-mode: compile

$h = hash('sha1', 'hello');
echo strlen($h) === 40 ? 'ok' : 'fail';
echo $h === sha1('hello') ? ':matches' : ':differs';
