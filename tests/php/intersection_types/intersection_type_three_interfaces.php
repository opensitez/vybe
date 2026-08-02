<?php
// vybe-test: php/intersection_types/intersection_type_three_interfaces
// origin: languages/php/tests/php/test_intersection_types.rs

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
interface C { public function c(): int; }
class ABC implements A, B, C {
    public function a(): int { return 1; }
    public function b(): int { return 2; }
    public function c(): int { return 3; }
}
function sum(A&B&C $obj): int { return $obj->a() + $obj->b() + $obj->c(); }
echo sum(new ABC());

__vybe_check(ob_get_clean(), "6");
