<?php
// vybe-test: php/hash_crypto/random_int_range
// origin: languages/php/tests/php/test_hash_crypto.rs
// vybe-test-mode: compile

$n = random_int(1, 100);
echo $n >= 1 && $n <= 100 ? 'in range' : 'out of range';
echo is_int($n) ? ':int' : ':not int';
