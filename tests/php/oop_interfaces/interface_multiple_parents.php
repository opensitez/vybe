<?php
// vybe-test: php/oop_interfaces/interface_multiple_parents
// origin: languages/php/tests/php/test_oop_interfaces.rs

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

interface A { public function a(): int; }
interface B { public function b(): int; }
interface C extends A, B {}
class Impl implements C {
    public function a(): int { return 1; }
    public function b(): int { return 2; }
}
$o = new Impl;
echo $o->a() + $o->b();

__vybe_check(ob_get_clean(), "3");
