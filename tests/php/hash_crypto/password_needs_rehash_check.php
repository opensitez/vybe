<?php
// vybe-test: php/hash_crypto/password_needs_rehash_check
// origin: languages/php/tests/php/test_hash_crypto.rs
// vybe-test-mode: compile

$hash = password_hash('pass', PASSWORD_BCRYPT, ['cost' => 10]);
$same = password_needs_rehash($hash, PASSWORD_BCRYPT, ['cost' => 10]);
$more = password_needs_rehash($hash, PASSWORD_BCRYPT, ['cost' => 14]);
echo is_bool($same) ? 'bool' : 'not bool';
echo $more ? ':needs rehash' : ':no rehash';
