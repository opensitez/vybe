<?php
// vybe-test: php/programs/insertion_sort_correctness
// origin: languages/php/tests/php/test_programs.rs

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

function insertionSort(array $arr): array {
    for ($i = 1; $i < count($arr); $i++) {
        $key = $arr[$i];
        $j = $i - 1;
        while ($j >= 0 && $arr[$j] > $key) { $arr[$j+1] = $arr[$j]; $j--; }
        $arr[$j+1] = $key;
    }
    return $arr;
}
echo implode(',', insertionSort([9,3,7,1,5])) . "\n";

__vybe_check(ob_get_clean(), "1,3,5,7,9");
