<?php
// vybe-test: php/array_column_object_properties/array_column_objects_private_prop_access
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

class PrivateUser {
    private int $id;
    private string $name;
    public function __construct(int $id, string $name) {
        $this->id = $id;
        $this->name = $name;
    }
}
$users = [new PrivateUser(1, 'Alice'), new PrivateUser(2, 'Bob')];
$ids = array_column($users, 'id');
echo $ids[0] . "|" . $ids[1];

__vybe_check(ob_get_clean(), "1|2");
