<?php
// vybe-test: php/php_serialization_unserialize_allowed_classes/test_php_serialize_unserialize_primitive_types
// origin: languages/php/tests/php/test_php_serialization_unserialize_allowed_classes.rs

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

$data = [
    "int" => 42,
    "float" => 3.14,
    "string" => "hello",
    "bool" => true,
    "null" => null,
    "arr" => [1, 2, 3]
];

$serialized = serialize($data);
$restored = unserialize($serialized);
echo $restored["string"] . " " . $restored["int"] . " arr_count=" . count($restored["arr"]);

__vybe_check(ob_get_clean(), "hello 42 arr_count=3");
