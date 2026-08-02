<?php
// vybe-test: php/array_column_advanced/array_chunk_preserve_keys_false_reindexes_from_zero
// origin: languages/php/tests/php/test_array_column_advanced.rs

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

$chunks = array_chunk(['a' => 1, 'b' => 2, 'c' => 3, 'd' => 4], 3);
echo count($chunks) . '|' . array_key_first($chunks[0]) . '|' . implode(',', $chunks[1]);

__vybe_check(ob_get_clean(), "2|0|3,4");
