<?php
// vybe-test: php/abstract_final_patterns/abstract_class_constructor_called_by_child
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

abstract class Vehicle {
    public function __construct(public readonly string $make) {}
    abstract public function type(): string;
}
class Car extends Vehicle {
    public function type(): string { return "car"; }
}
$c = new Car("Toyota");
echo $c->make . ':' . $c->type(), "\n";

__vybe_check(ob_get_clean(), "Toyota:car");
