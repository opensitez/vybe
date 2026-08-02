<?php
// vybe-test: php/hash_crypto/password_get_info_structure
// origin: languages/php/tests/php/test_hash_crypto.rs
// vybe-test-mode: compile

$hash = password_hash('test', PASSWORD_BCRYPT);
$info = password_get_info($hash);
echo array_key_exists('algo', $info) ? 'has algo' : 'no algo';
echo array_key_exists('algoName', $info) ? ':has name' : ':no name';
echo array_key_exists('options', $info) ? ':has opts' : ':no opts';
