<?php
// vybe-test: php/password_security/password_get_info_bcrypt
// origin: languages/php/tests/php/test_password_security.rs
// vybe-test-mode: compile

$hash = password_hash('test', PASSWORD_BCRYPT);
$info = password_get_info($hash);
echo isset($info['algo']) ? 'has algo' : 'no algo';
echo isset($info['algoName']) ? ':has algoName' : ':no algoName';
echo isset($info['options']) ? ':has options' : ':no options';
