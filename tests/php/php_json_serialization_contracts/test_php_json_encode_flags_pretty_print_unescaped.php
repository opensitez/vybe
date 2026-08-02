<?php
// vybe-test: php/php_json_serialization_contracts/test_php_json_encode_flags_pretty_print_unescaped
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

$data = ["url" => "https://example.com/api", "title" => "Home & About"];
$json = json_encode($data, JSON_UNESCAPED_SLASHES | JSON_UNESCAPED_UNICODE);
echo $json;

__vybe_check(ob_get_clean(), "{\"url\":\"https://example.com/api\",\"title\":\"Home & About\"}");
