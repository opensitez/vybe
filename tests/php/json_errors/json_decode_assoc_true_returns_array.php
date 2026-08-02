<?php
// vybe-test: php/json_errors/json_decode_assoc_true_returns_array
// origin: languages/php/tests/php/test_json_errors.rs

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

$data = json_decode('{"k":2}', true, 512, JSON_THROW_ON_ERROR);
echo is_array($data) && $data['k'] === 2 ? 'assoc' : 'no';

__vybe_check(ob_get_clean(), "assoc");
