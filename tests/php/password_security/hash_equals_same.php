<?php
// vybe-test: php/password_security/hash_equals_same
// origin: languages/php/tests/php/test_password_security.rs
// vybe-test-mode: compile

$a = hash('sha256', 'hello');
$b = hash('sha256', 'hello');
echo hash_equals($a, $b) ? 'equal' : 'not equal';
