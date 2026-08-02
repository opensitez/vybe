<?php
// vybe-test: php/json_errors/json_encode_max_depth_nested_arrays
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

$a = []; $cursor = &$a;
for ($i = 0; $i < 3; $i++) { $cursor['n'] = []; $cursor = &$cursor['n']; }
echo json_encode($a, JSON_THROW_ON_ERROR);

__vybe_check(ob_get_clean(), "{\"n\":{\"n\":{\"n\":[]}}}");
