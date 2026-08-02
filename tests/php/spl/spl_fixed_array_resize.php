<?php
// vybe-test: php/spl/spl_fixed_array_resize
// origin: languages/php/tests/php/test_spl.rs
// vybe-test-mode: compile

$arr = new SplFixedArray(3);
$arr[0] = 'a'; $arr[1] = 'b'; $arr[2] = 'c';
$arr->setSize(5);
$arr[3] = 'd'; $arr[4] = 'e';
echo $arr->getSize() . ':' . $arr[4];
