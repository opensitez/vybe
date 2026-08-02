<?php
// vybe-test: php/strings/string_interpolation_nested_index_runtime
// origin: languages/php/tests/php/test_strings.rs

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

$payload = ['user' => ['name' => 'Alice', 'roles' => ['admin', 'editor']]];
echo "{$payload['user']['name']}=>{$payload['user']['roles'][1]}";

__vybe_check(ob_get_clean(), "Alice=>editor");
