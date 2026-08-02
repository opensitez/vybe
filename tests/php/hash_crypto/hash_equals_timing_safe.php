<?php
// vybe-test: php/hash_crypto/hash_equals_timing_safe
// origin: languages/php/tests/php/test_hash_crypto.rs
// vybe-test-mode: compile

$a = hash('sha256', 'secret');
$b = hash('sha256', 'secret');
$c = hash('sha256', 'other');
echo hash_equals($a, $b) ? 'equal' : 'not equal';
echo hash_equals($a, $c) ? 'equal' : 'not equal';
