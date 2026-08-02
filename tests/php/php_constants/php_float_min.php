<?php
// vybe-test: php/php_constants/php_float_min
// origin: languages/php/tests/php/test_php_constants.rs
// vybe-test-mode: compile

$x = PHP_FLOAT_MIN;
echo ($x > 0) ? 'positive' : 'unexpected';
