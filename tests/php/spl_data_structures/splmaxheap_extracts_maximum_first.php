<?php
// vybe-test: php/spl_data_structures/splmaxheap_extracts_maximum_first
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

$h = new SplMaxHeap;
$h->insert(5); $h->insert(2); $h->insert(8); $h->insert(1);
$out = [];
while (!$h->isEmpty()) $out[] = $h->extract();
echo implode(',', $out);

__vybe_check(ob_get_clean(), "8,5,2,1");
