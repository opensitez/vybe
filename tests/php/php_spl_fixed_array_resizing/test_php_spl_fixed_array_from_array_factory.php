<?php
// vybe-test: php/php_spl_fixed_array_resizing/test_php_spl_fixed_array_from_array_factory
// origin: languages/php/tests/php/test_php_spl_fixed_array_resizing.rs
// vybe-test-mode: compile

$native = ["foo" => 10, "bar" => 20];
$fixed = SplFixedArray::fromArray($native, false);
echo $fixed->getSize() === 2 && $fixed[0] === 10 ? "FROM_ARRAY_OK" : "FAIL";
