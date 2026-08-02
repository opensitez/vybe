<?php
// vybe-test: php/type_checks_extra/print_r_return_true_assign
// origin: languages/php/tests/php/test_type_checks_extra.rs
// vybe-test-mode: compile

$arr = ['a' => 1, 'b' => 2];
$out = print_r($arr, true);
echo strlen($out) > 0 ? 'has output' : 'empty';
