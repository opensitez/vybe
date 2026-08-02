<?php
// vybe-test: php/php_spl_data_structures_fixed_array/test_php_spl_fixed_array_resize
// origin: languages/php/tests/php/test_php_spl_data_structures_fixed_array.rs
// vybe-test-mode: compile

$fa = new SplFixedArray(2);
$fa[0] = "a";
$fa->setSize(4);
$fa[3] = "d";
echo $fa->getSize() . " | " . $fa[3];
