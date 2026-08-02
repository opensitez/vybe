<?php
// vybe-test: php/php_spl_fixed_array_resizing/test_php_spl_fixed_array_to_array_export
// origin: languages/php/tests/php/test_php_spl_fixed_array_resizing.rs
// vybe-test-mode: compile

$fixed = new SplFixedArray(2);
$fixed[0] = "x";
$fixed[1] = "y";
$exported = $fixed->toArray();
echo is_array($exported) && $exported[0] === "x" ? "TO_ARRAY_OK" : "FAIL";
