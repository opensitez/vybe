<?php
// vybe-test: php/php_constants/sort_flag_constants
// origin: languages/php/tests/php/test_php_constants.rs
// vybe-test-mode: compile

$a = [3, 1, 2];
sort($a, SORT_REGULAR);
$b = ['10', '9', '2'];
sort($b, SORT_NUMERIC);
$c = ['banana', 'apple', 'cherry'];
sort($c, SORT_STRING);
echo count($a) + count($b) + count($c);
