<?php
// vybe-test: php/union_types/contravariant_param_in_override
// origin: languages/php/tests/php/test_union_types.rs

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

class AnimalFood {}
class DogFood extends AnimalFood {}
interface Handler { public function handle(DogFood $f): void; }
class AnyHandler implements Handler { public function handle(AnimalFood $f): void { echo get_class($f); } }
(new AnyHandler)->handle(new DogFood);

__vybe_check(ob_get_clean(), "DogFood");
