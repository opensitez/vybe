<?php
// vybe-test: php/operators/instanceof_basic_runtime_operator
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

class Base {}
class Child extends Base {}
$base = new Base();
$child = new Child();
echo $base instanceof Base ? 'b1' : 'b0';
echo '|';
echo $child instanceof Base ? 'c1' : 'c0';
echo '|';
echo $base instanceof Child ? 'd1' : 'd0';

__vybe_check(ob_get_clean(), "b1|c1|d0");
