<?php
// vybe-test: php/hash_crypto/hash_algos_list
// origin: languages/php/tests/php/test_hash_crypto.rs
// vybe-test-mode: compile

$algos = hash_algos();
echo is_array($algos) ? 'array' : 'not array';
echo in_array('sha256', $algos) ? ':sha256' : ':no sha256';
echo in_array('md5', $algos) ? ':md5' : ':no md5';
