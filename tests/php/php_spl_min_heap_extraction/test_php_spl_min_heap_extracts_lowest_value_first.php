<?php
// vybe-test: php/php_spl_min_heap_extraction/test_php_spl_min_heap_extracts_lowest_value_first
// origin: languages/php/tests/php/test_php_spl_min_heap_extraction.rs

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

$heap = new SplMinHeap();
$heap->insert(50);
$heap->insert(10);
$heap->insert(30);

$items = [];
while ($heap->valid()) {
    $items[] = $heap->extract();
}
echo implode(",", $items);

__vybe_check(ob_get_clean(), "10,30,50");
