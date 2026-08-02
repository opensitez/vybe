<?php
// vybe-test: php/array_creation/array_change_key_case_upper
// origin: languages/php/tests/php/test_array_creation.rs

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

echo json_encode(array_change_key_case(['foo' => 1, 'Bar' => 2], CASE_UPPER));

__vybe_check(ob_get_clean(), "{\"FOO\":1,\"BAR\":2}");
