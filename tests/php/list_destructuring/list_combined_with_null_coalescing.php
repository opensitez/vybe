<?php
// vybe-test: php/list_destructuring/list_combined_with_null_coalescing
// origin: languages/php/tests/php/test_list_destructuring.rs

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

$response = ['status' => 200, 'body' => 'ok'];
['status' => $code, 'headers' => $hdrs] = $response + ['headers' => []];
echo $code . ':' . count($hdrs);

__vybe_check(ob_get_clean(), "200:0");
