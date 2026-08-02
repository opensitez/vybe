<?php
// vybe-test: php/oop/reflection_property_methods_counts_runtime
// origin: languages/php/tests/php/test_oop.rs

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

class User {
    public int $id;
    private string $name;
    public function __construct() {}
    public function hello(): string { return 'hi'; }
}
$obj = new User();
echo property_exists($obj, 'id') ? 'id' : 'no';
echo method_exists($obj, 'hello') ? '|hello' : '|no';

__vybe_check(ob_get_clean(), "id|hello");
