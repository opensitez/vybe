<?php
// vybe-test: php/inheritance_patterns/is_subclass_of
// origin: languages/php/tests/php/test_inheritance_patterns.rs

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

class Animal {} class Mammal extends Animal {} class Dog extends Mammal {}
echo is_subclass_of(Dog::class, Animal::class) ? 'yes' : 'no', "\n";
echo is_subclass_of(Animal::class, Dog::class) ? 'yes' : 'no', "\n";

__vybe_check(ob_get_clean(), "yes\nno");
