<?php
// vybe-test: php/spl/spl_fixed_array_basic
// origin: languages/php/tests/php/test_spl.rs
// vybe-test-mode: compile

$arr = new SplFixedArray(5);
$arr[0] = 10;
$arr[1] = 20;
$arr[2] = 30;
echo $arr->getSize() . ':' . $arr[1];
