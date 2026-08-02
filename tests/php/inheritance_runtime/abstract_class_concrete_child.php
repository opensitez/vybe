<?php
// vybe-test: php/inheritance_runtime/abstract_class_concrete_child
// origin: languages/php/tests/php/test_inheritance_runtime.rs

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

abstract class A { abstract public function f(): int; }
class C extends A { public function f(): int { return 3; } }
echo (new C())->f();

__vybe_check(ob_get_clean(), "3");
