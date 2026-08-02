<?php
// vybe-test: php/php_math_random_engines_php82/test_php_mt_rand_mt_srand_seed_reproducibility
// origin: languages/php/tests/php/test_php_math_random_engines_php82.rs
// vybe-test-mode: compile

mt_srand(12345, MT_RAND_MT19937);
$v1 = mt_rand(1, 100);
mt_srand(12345, MT_RAND_MT19937);
$v2 = mt_rand(1, 100);
echo $v1 === $v2 ? "MT_REPRODUCIBLE" : "FAIL";
