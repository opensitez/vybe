<?php
// vybe-test: php/array_search_sort/array_search_offset_starts_after_index
// origin: languages/php/tests/php/test_array_search_sort.rs

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

// array_search() takes at most 3 arguments — a 4th "offset" is an
// ArgumentCountError, not a search start. Starting the search after an index is
// spelled with a key-preserving array_slice.
$a = ['2', 3, '2', 4];
echo array_search('2', array_slice($a, 2, null, true), false);

__vybe_check(ob_get_clean(), "2");
