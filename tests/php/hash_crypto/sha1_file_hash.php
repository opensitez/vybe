<?php
// vybe-test: php/hash_crypto/sha1_file_hash
// origin: languages/php/tests/php/test_hash_crypto.rs
// vybe-test-mode: compile

$result = sha1_file('/etc/hostname');
echo is_string($result) || $result === false ? 'ok' : 'fail';
if (is_string($result)) {
    echo strlen($result) === 40 ? ':len ok' : ':len fail';
}
