<?php
// vybe-test: php/references/pass_by_reference_nested_array
// origin: languages/php/tests/php/test_references.rs
// vybe-test-mode: compile

function doubleAll(array &$arr) {
    foreach ($arr as &$v) { $v *= 2; }
}
$data = [1, 2, 3, 4];
doubleAll($data);
echo implode(',', $data);
