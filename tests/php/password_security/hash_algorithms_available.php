<?php
// vybe-test: php/password_security/hash_algorithms_available
// origin: languages/php/tests/php/test_password_security.rs
// vybe-test-mode: compile

$algos = hash_algos();
echo in_array('sha256', $algos) ? 'sha256 available' : 'missing sha256';
echo in_array('sha512', $algos) ? ':sha512 available' : ':missing sha512';
echo in_array('md5',    $algos) ? ':md5 available'    : ':missing md5';
