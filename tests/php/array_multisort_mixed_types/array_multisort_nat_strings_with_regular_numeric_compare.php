<?php
// vybe-test: php/array_multisort_mixed_types/array_multisort_nat_strings_with_regular_numeric_compare
// origin: languages/php/tests/php/test_array_multisort_mixed_types.rs

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

$versions = ["v2", "v10", "v1", "v3"];
$nums = [1, 2, 3, 4];
array_multisort($versions, SORT_NATURAL | SORT_FLAG_CASE, SORT_NUMERIC, $nums, SORT_DESC, SORT_NUMERIC);
echo implode(',', $versions) . "|" . implode(',', $nums);

__vybe_check(ob_get_clean(), "v1,v2,v3,v10|3,1,4,2");
