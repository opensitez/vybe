<?php
// vybe-test: php/references/return_by_reference_function
// origin: languages/php/tests/php/test_references.rs
// vybe-test-mode: compile

function &getRef(array &$arr, $key) {
    return $arr[$key];
}
$data = ['a' => 1];
$ref = &getRef($data, 'a');
$ref = 99;
echo $data['a'];
