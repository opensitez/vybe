<?php
// vybe-test: php/abstract_final_patterns/abstract_method_final_in_intermediate_class
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

abstract class A2 { abstract public function run(): string; }
class B2 extends A2 { final public function run(): string { return "B2"; } }
class C2 extends B2 {}
echo (new C2())->run(), "\n";

__vybe_check(ob_get_clean(), "B2");
