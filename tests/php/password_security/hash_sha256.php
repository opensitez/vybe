<?php
// vybe-test: php/password_security/hash_sha256
// origin: languages/php/tests/php/test_password_security.rs
// vybe-test-mode: compile

$h = hash('sha256', 'hello');
echo strlen($h) === 64 ? 'sha256 length ok' : 'wrong length';
echo ctype_xdigit($h) ? ':hex chars' : ':not hex';
