<?php
// vybe-test: php/php_constants/php_float_max
// origin: languages/php/tests/php/test_php_constants.rs
// vybe-test-mode: compile

$x = PHP_FLOAT_MAX;
echo ($x > 0) ? 'positive' : 'unexpected';
