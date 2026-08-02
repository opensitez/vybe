<?php
// vybe-test: php/php_spl_fixed_array_resizing/test_php_spl_fixed_array_shrink_truncates_elements
// origin: languages/php/tests/php/test_php_spl_fixed_array_resizing.rs
// vybe-test-mode: compile

$fixed = new SplFixedArray(5);
$fixed[4] = "last";
$fixed->setSize(2);
echo $fixed->getSize() === 2 ? "SHRINK_OK" : "FAIL";
