<?php
// vybe-test: php/oop_patterns/covariant_return_type_runtime
// origin: languages/php/tests/php/test_oop_patterns.rs

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

class Animal {}
class Dog extends Animal {}
class AnimalFactory { public function create(): Animal { return new Animal(); } }
class DogFactory extends AnimalFactory {
    public function create(): Dog { return new Dog(); }
}
$f = new DogFactory();
echo $f->create() instanceof Dog ? 'yes' : 'no';

__vybe_check(ob_get_clean(), "yes");
