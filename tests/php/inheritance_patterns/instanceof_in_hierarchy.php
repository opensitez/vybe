<?php
// vybe-test: php/inheritance_patterns/instanceof_in_hierarchy
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

class Vehicle {}
class Car extends Vehicle {}
class ElectricCar extends Car {}
$e = new ElectricCar;
echo ($e instanceof ElectricCar ? '1' : '0') . ($e instanceof Car ? '1' : '0') . ($e instanceof Vehicle ? '1' : '0'), "\n";

__vybe_check(ob_get_clean(), "111");
