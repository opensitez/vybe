<?php
// vybe-test: php/advanced_oop/constructor_visibility_chain
// origin: languages/php/tests/php/test_advanced_oop.rs

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
    public function __construct(public string $name) {}
}
class Child extends Base {
    public function __construct(string $name, public int $id) {
        parent::__construct($name);
    }
}
$child = new Child('worker', 7);
echo $child->name . ':' . $child->id;

__vybe_check(ob_get_clean(), "worker:7");
