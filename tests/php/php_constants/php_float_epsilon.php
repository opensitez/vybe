<?php
// vybe-test: php/php_constants/php_float_epsilon
// origin: languages/php/tests/php/test_php_constants.rs
// vybe-test-mode: compile

$eps = PHP_FLOAT_EPSILON;
echo ($eps > 0) ? 'positive' : 'unexpected';
