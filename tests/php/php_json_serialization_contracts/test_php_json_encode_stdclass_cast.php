<?php
// vybe-test: php/php_json_serialization_contracts/test_php_json_encode_stdclass_cast
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

$obj = (object)["key" => "value", "id" => 123];
echo json_encode($obj);

__vybe_check(ob_get_clean(), "{\"key\":\"value\",\"id\":123}");
