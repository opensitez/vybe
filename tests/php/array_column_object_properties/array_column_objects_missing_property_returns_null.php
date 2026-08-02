<?php
// vybe-test: php/array_column_object_properties/array_column_objects_missing_property_returns_null
// origin: languages/php/tests/php/test_array_column_object_properties.rs

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

class Basic {
    public function __construct(public int $id) {}
}
$rows = [new Basic(1)];
$vals = array_column($rows, 'does_not_exist');
echo $vals[0] === null ? 'null' : 'notnull';

__vybe_check(ob_get_clean(), "null");
