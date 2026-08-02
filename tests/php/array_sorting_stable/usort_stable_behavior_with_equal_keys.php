<?php
// vybe-test: php/array_sorting_stable/usort_stable_behavior_with_equal_keys
// origin: languages/php/tests/php/test_array_sorting_stable.rs

function __vybe_check($got, $want) {
    // Match the Rust harness's normalisation: strip \r, then drop trailing
    // newlines (it split on "\n" and popped empty trailing elements).
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    // Replay the program's own output so running the file by hand still
    // behaves like the program it was extracted from.
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

$rows = [
    ['id' => 1, 'group' => 'a'],
    ['id' => 2, 'group' => 'a'],
    ['id' => 3, 'group' => 'b'],
];
usort($rows, fn($x, $y) => strcmp($x['group'], $y['group']));
echo $rows[0]['id'] . ':' . $rows[1]['id'] . ':' . $rows[2]['id'];

__vybe_check(ob_get_clean(), "1:2:3");
