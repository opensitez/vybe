<?php
// vybe-test: php/php_spl_data_structures_fixed_array/test_php_spl_fixed_array_from_array_conversion
// origin: languages/php/tests/php/test_php_spl_data_structures_fixed_array.rs
// vybe-test-mode: compile

$native = [100, 200, 300];
$fixed = SplFixedArray::fromArray($native);
echo $fixed->getSize();
print_r($fixed->toArray());
