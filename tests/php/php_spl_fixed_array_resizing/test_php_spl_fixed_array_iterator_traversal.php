<?php
// vybe-test: php/php_spl_fixed_array_resizing/test_php_spl_fixed_array_iterator_traversal
// origin: languages/php/tests/php/test_php_spl_fixed_array_resizing.rs
// vybe-test-mode: compile

$fixed = new SplFixedArray(3);
$fixed[0] = 10; $fixed[1] = 20; $fixed[2] = 30;
$sum = 0;
foreach ($fixed as $val) {
    $sum += $val;
}
echo $sum === 60 ? "FIXED_ITERATOR_OK" : "FAIL";
