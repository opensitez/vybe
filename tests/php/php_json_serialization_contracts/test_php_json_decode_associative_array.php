<?php
// vybe-test: php/php_json_serialization_contracts/test_php_json_decode_associative_array
// origin: languages/php/tests/php/test_php_json_serialization_contracts.rs

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

$json = '{"name":"Alice","skills":["PHP","Rust"]}';
$data = json_decode($json, true);
echo $data["name"] . " -> " . implode(",", $data["skills"]);

__vybe_check(ob_get_clean(), "Alice -> PHP,Rust");
