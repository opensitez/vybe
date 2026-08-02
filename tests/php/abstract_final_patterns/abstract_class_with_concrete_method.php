<?php
// vybe-test: php/abstract_final_patterns/abstract_class_with_concrete_method
// origin: languages/php/tests/php/test_abstract_final_patterns.rs

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

abstract class Base {
    public function identify(): string { return "base"; }
    abstract public function tag(): string;
}
class Child extends Base {
    public function tag(): string { return "child"; }
}
$c = new Child();
echo $c->identify() . ',' . $c->tag(), "\n";

__vybe_check(ob_get_clean(), "base,child");
