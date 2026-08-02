<?php
// vybe-test: php/array_column_object_properties/array_column_objects_non_scalar_index_uses_column_stringification
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

class Row {
    public function __construct(public int $id, public string $name) {}
}
$rows = [new Row(1, 'A'), new Row(2, 'B')];
$vals = array_column($rows, 'name', 'id');
echo $vals[1] . '|' . $vals[2];

__vybe_check(ob_get_clean(), "A|B");
