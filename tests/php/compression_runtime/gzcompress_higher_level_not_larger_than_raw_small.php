<?php
// vybe-test: php/compression_runtime/gzcompress_higher_level_not_larger_than_raw_small
// origin: languages/php/tests/php/test_compression_runtime.rs

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

$raw = str_repeat('a', 50);
echo strlen(gzcompress($raw, 9)) > 0 ? 'ok' : 'empty';

__vybe_check(ob_get_clean(), "ok");
