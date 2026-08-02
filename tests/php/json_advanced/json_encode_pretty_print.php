<?php
// vybe-test: php/json_advanced/json_encode_pretty_print
// origin: languages/php/tests/php/test_json_advanced.rs

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

$s = json_encode(['a'=>1,'b'=>2], JSON_PRETTY_PRINT);
echo str_contains($s, "\n") ? 'multiline' : 'single';

__vybe_check(ob_get_clean(), "multiline");
