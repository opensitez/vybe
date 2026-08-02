<?php
// vybe-test: php/password_security/password_bcrypt_cost
// origin: languages/php/tests/php/test_password_security.rs
// vybe-test-mode: compile

$hash = password_hash('test', PASSWORD_BCRYPT, ['cost' => 4]);
$info = password_get_info($hash);
echo $info['options']['cost'] === 4 ? 'cost 4' : 'different cost';
