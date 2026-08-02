<?php
// vybe-test: php/programs/bubble_sort_correctness
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

function bubbleSort(array $arr): array {
    $n = count($arr);
    for ($i = 0; $i < $n; $i++)
        for ($j = 0; $j < $n - $i - 1; $j++)
            if ($arr[$j] > $arr[$j+1]) { $t=$arr[$j]; $arr[$j]=$arr[$j+1]; $arr[$j+1]=$t; }
    return $arr;
}
echo implode(',', bubbleSort([5,3,1,4,2])) . "\n";

__vybe_check(ob_get_clean(), "1,2,3,4,5");
