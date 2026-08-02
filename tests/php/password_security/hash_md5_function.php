<?php
// vybe-test: php/password_security/hash_md5_function
// origin: languages/php/tests/php/test_password_security.rs
// vybe-test-mode: compile

$h = hash('md5', 'hello');
echo strlen($h) === 32 ? 'md5 length ok' : 'wrong length';
echo $h === md5('hello') ? ':matches md5()' : ':different';
