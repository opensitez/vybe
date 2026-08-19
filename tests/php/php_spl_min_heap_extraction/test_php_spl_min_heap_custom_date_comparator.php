<?php
// vybe-test: php/php_spl_min_heap_extraction/test_php_spl_min_heap_custom_date_comparator
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

class DateMinHeap extends SplMinHeap {
    protected function compare(mixed $val1, mixed $val2): int {
        return strtotime($val1) <=> strtotime($val2);
    }
}

$heap = new DateMinHeap();
$heap->insert("2024-12-31");
$heap->insert("2024-01-01");
$heap->insert("2024-06-15");

echo "Earliest date: " . $heap->extract();

__vybe_check(ob_get_clean(), "Earliest date: 2024-12-31");
