<?php
// vybe-test: php/php_spl_fixed_array_resizing/test_php_spl_fixed_array_offset_unset_resets_to_null
// origin: languages/php/tests/php/test_php_spl_fixed_array_resizing.rs
// vybe-test-mode: compile

$fixed = new SplFixedArray(2);
$fixed[0] = "data";
unset($fixed[0]);
echo $fixed[0] === null ? "UNSET_NULL_OK" : "FAIL";
