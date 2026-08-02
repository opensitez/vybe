<?php
// vybe-test: php/array_sorting_stable/arsort_with_ties_still_stable
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

$a = ['first' => 1, 'second' => 1, 'third' => 2];
arsort($a);
echo array_key_first($a) . '|' . array_key_last($a);

__vybe_check(ob_get_clean(), "third|second");
