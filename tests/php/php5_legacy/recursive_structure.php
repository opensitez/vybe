<?php
// vybe-test: php/php5_legacy/recursive_structure
// origin: languages/php/tests/php/test_php5_legacy.rs
// vybe-test-mode: compile

function flatten($arr) {
    $result = [];
    foreach ($arr as $item) {
        if (is_array($item)) {
            $sub = flatten($item);
            $result = array_merge($result, $sub);
        } else {
            array_push($result, $item);
        }
    }
    return $result;
}
$nested = [1, [2, 3], [4, [5, 6]]];
$flat = flatten($nested);
