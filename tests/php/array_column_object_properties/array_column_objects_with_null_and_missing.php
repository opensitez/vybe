<?php
// vybe-test: php/array_column_object_properties/array_column_objects_with_null_and_missing
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

class SoftUser {
    public function __construct(public int $id, public ?string $name = null) {}
}
$rows = [new SoftUser(1, 'Alice'), new SoftUser(2), new SoftUser(3, 'Bob')];
$names = array_column($rows, 'name');
echo implode('|', $names);

__vybe_check(ob_get_clean(), "Alice||Bob");
