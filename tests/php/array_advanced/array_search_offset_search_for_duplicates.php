<?php
// vybe-test: php/array_advanced/array_search_offset_search_for_duplicates
// origin: languages/php/tests/php/test_array_advanced.rs

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

$values = ["zero", "needle", "skip", "needle", "end"];
echo array_search("needle", $values, true);
echo "|";
// array_search() has no 4th "offset" parameter (ArgumentCountError); resuming
// the search past an index is a key-preserving array_slice.
echo array_search("needle", array_slice($values, 3, null, true), true);

__vybe_check(ob_get_clean(), "1|3");
