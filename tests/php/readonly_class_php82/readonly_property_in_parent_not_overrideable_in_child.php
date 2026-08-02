<?php
// vybe-test: php/readonly_class_php82/readonly_property_in_parent_not_overrideable_in_child
// origin: languages/php/tests/php/test_readonly_class_php82.rs

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

class Vehicle {
    public readonly string $make;
    public function __construct(string $make) { $this->make = $make; }
}
class Car extends Vehicle {
    public function __construct(string $make, public readonly int $year) {
        parent::__construct($make);
    }
}
$c = new Car("Toyota", 2020);
echo $c->make . ',' . $c->year;

__vybe_check(ob_get_clean(), "Toyota,2020");
