<?php
// vybe-test: php/covariant_return_types/contravariant_parameter_widens_type
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

class Animal { public function name(): string { return "animal"; } }
class Dog extends Animal { public function name(): string { return "dog"; } }
interface Feeder { public function feed(Dog $dog): void; }
class GenericFeeder implements Feeder {
    public function feed(Animal $animal): void { echo "feeding " . $animal->name(); }
}
$feeder = new GenericFeeder();
$feeder->feed(new Dog());

__vybe_check(ob_get_clean(), "feeding dog");
