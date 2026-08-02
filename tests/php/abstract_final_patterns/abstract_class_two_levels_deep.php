<?php
// vybe-test: php/abstract_final_patterns/abstract_class_two_levels_deep
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

abstract class A { abstract public function name(): string; }
abstract class B extends A { abstract public function age(): int; }
class C extends B {
    public function name(): string { return "Carol"; }
    public function age(): int { return 25; }
}
$c = new C();
echo $c->name() . ',' . $c->age(), "\n";

__vybe_check(ob_get_clean(), "Carol,25");
