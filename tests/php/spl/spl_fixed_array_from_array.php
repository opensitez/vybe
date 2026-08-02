<?php
// vybe-test: php/spl/spl_fixed_array_from_array
// origin: languages/php/tests/php/test_spl.rs
// vybe-test-mode: compile

$arr = SplFixedArray::fromArray([100, 200, 300]);
echo $arr->getSize() . ':' . $arr[2];
