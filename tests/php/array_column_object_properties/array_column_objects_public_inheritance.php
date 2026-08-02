<?php
// vybe-test: php/array_column_object_properties/array_column_objects_public_inheritance
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

class Base {
    public function __construct(public int $id) {}
}
class Child extends Base {
    public function __construct(int $id, public string $name) {
        parent::__construct($id);
    }
}
$rows = [new Child(1, 'Ada'), new Child(2, 'Lin')];
$vals = array_column($rows, 'name');
echo implode('|', $vals);

__vybe_check(ob_get_clean(), "Ada|Lin");
