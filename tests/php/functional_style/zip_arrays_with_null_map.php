<?php
// vybe-test: php/functional_style/zip_arrays_with_null_map
// origin: languages/php/tests/php/test_functional_style.rs
// vybe-test-mode: compile

$keys   = ['a', 'b', 'c'];
$values = [1,   2,   3  ];
$zipped = array_map(null, $keys, $values);
foreach ($zipped as [$k, $v]) {
    echo "$k=$v ";
}
