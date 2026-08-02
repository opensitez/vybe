<?php
// vybe-test: php/json_encode/json_encode_unescaped_slashes
// origin: languages/php/tests/php/test_json_encode.rs

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

echo json_encode(['u' => 'https://ex.com/a/b'], JSON_UNESCAPED_SLASHES);

__vybe_check(ob_get_clean(), "{\"u\":\"https://ex.com/a/b\"}");
