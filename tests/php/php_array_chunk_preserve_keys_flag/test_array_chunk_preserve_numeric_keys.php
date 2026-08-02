<?php
// vybe-test: php/php_array_chunk_preserve_keys_flag/test_array_chunk_preserve_numeric_keys
// origin: languages/php/tests/php/test_php_array_chunk_preserve_keys_flag.rs

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

$input = [10 => 'a', 20 => 'b', 30 => 'c'];
$chunks = array_chunk($input, 2, true);
echo implode(',', array_keys($chunks[0])) . '|' . implode(',', array_keys($chunks[1])), "\n";

__vybe_check(ob_get_clean(), "10,20|30");
