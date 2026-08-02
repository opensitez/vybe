<?php
// vybe-test: php/array_column_advanced/array_splice_replaces_with_more_items_and_return_count
// origin: languages/php/tests/php/test_array_column_advanced.rs

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

$items = [1, 2, 3, 4];
$removed = array_splice($items, 1, 1, [9, 9, 9]);
echo count($removed) . '|' . implode(',', $items);

__vybe_check(ob_get_clean(), "1|1,9,9,9,3,4");
