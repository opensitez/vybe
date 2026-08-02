<?php
// vybe-test: php/password_security/hash_sha512
// origin: languages/php/tests/php/test_password_security.rs
// vybe-test-mode: compile

$h = hash('sha512', 'hello');
echo strlen($h) === 128 ? 'sha512 length ok' : 'wrong length';
