<?php
// vybe-test: php/spl_data_structures/splminheap_count
// origin: languages/php/tests/php/test_spl_data_structures.rs

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

$h = new SplMinHeap;
$h->insert(3); $h->insert(1);
echo count($h);

__vybe_check(ob_get_clean(), "2");
