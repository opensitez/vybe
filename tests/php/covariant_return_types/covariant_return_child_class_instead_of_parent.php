<?php
// vybe-test: php/covariant_return_types/covariant_return_child_class_instead_of_parent
// origin: languages/php/tests/php/test_covariant_return_types.rs

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

class Shape { public function describe(): string { return "shape"; } }
class Circle extends Shape { public function describe(): string { return "circle"; } }
class ShapeFactory {
    public function make(): Shape { return new Shape(); }
}
class CircleFactory extends ShapeFactory {
    public function make(): Circle { return new Circle(); }
}
$factory = new CircleFactory();
echo $factory->make()->describe();

__vybe_check(ob_get_clean(), "circle");
