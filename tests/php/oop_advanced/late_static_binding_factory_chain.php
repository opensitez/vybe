<?php
// vybe-test: php/oop_advanced/late_static_binding_factory_chain
// origin: languages/php/tests/php/test_oop_advanced.rs

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
    protected string $color = "white";
    public static function make(): static {
        return new static();
    }
    public function paint(string $c): static {
        $clone = clone $this;
        $clone->color = $c;
        return $clone;
    }
    public function describe(): string {
        return static::class . ":" . $this->color;
    }
}
class Car extends Vehicle {}
$v = Vehicle::make()->paint("red");
$c = Car::make()->paint("blue");
echo $v->describe(), "\n";
echo $c->describe(), "\n";

__vybe_check(ob_get_clean(), "Vehicle:red\nCar:blue");
