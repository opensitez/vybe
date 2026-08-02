<?php
// vybe-test: php/hash_crypto/password_verify_correct_and_wrong
// origin: languages/php/tests/php/test_hash_crypto.rs
// vybe-test-mode: compile

$hash = password_hash('correct', PASSWORD_DEFAULT);
echo password_verify('correct', $hash) ? 'ok' : 'fail';
echo password_verify('wrong', $hash) ? 'ok' : ':rejected';
