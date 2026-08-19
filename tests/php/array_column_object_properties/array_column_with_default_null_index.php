<?php
// vybe-test: php/array_column_object_properties/array_column_with_default_null_index
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
    public function __construct(public int $id, public ?string $name = null) {}
}
$rows = [new Row(1, 'A'), new Row(2, null)];
$names = array_column($rows, 'name', 'missing');
echo json_encode(array_values($names));

__vybe_check(ob_get_clean(), "[\"A\",null]");
