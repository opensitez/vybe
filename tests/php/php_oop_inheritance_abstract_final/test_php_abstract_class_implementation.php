<?php
// vybe-test: php/php_oop_inheritance_abstract_final/test_php_abstract_class_implementation
// origin: languages/php/tests/php/test_php_oop_inheritance_abstract_final.rs

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

abstract class Shape {
    abstract public function area(): float;
}

class Circle extends Shape {
    public function __construct(public float $radius) {}
    public function area(): float {
        return 3.14159 * $this->radius * $this->radius;
    }
}

$c = new Circle(2.0);
echo round($c->area(), 2);

__vybe_check(ob_get_clean(), "12.57");
