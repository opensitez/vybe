<?php
// vybe-test: php/classes/class_instanceof_self_parent_runtime
// origin: languages/php/tests/php/test_classes.rs

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

class Vehicle { public function isVehicle(self $v): bool { return $v instanceof self; } }
class Car extends Vehicle {}
echo (new Vehicle())->isVehicle(new Vehicle()) ? 'yes' : 'no';
echo '|';
echo (new Vehicle())->isVehicle(new Car()) ? 'yes' : 'no';

__vybe_check(ob_get_clean(), "yes|no");
