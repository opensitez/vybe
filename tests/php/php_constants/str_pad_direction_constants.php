<?php
// vybe-test: php/php_constants/str_pad_direction_constants
// origin: languages/php/tests/php/test_php_constants.rs
// vybe-test-mode: compile

$left  = str_pad('5', 3, '0', STR_PAD_LEFT);
$right = str_pad('5', 3, '0', STR_PAD_RIGHT);
$both  = str_pad('5', 5, '-', STR_PAD_BOTH);
echo $left . $right . $both;
