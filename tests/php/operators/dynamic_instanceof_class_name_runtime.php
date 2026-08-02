<?php
// vybe-test: php/operators/dynamic_instanceof_class_name_runtime
// origin: languages/php/tests/php/test_operators.rs

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

class ParentClass {}
class ChildClass extends ParentClass {}
$obj = new ChildClass();
$type = ChildClass::class;
echo ($obj instanceof $type) . '|';
$type = ParentClass::class;
echo ($obj instanceof $type) . '|';
$type = stdClass::class;
echo ($obj instanceof $type ? 'yes' : 'no');

__vybe_check(ob_get_clean(), "1|1|no");
