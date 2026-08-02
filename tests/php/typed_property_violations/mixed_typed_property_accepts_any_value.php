<?php
// vybe-test: php/typed_property_violations/mixed_typed_property_accepts_any_value
// origin: languages/php/tests/php/test_typed_property_violations.rs

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

class Any { public mixed $slot; }
$a = new Any();
$a->slot = new ArrayObject([1]);
echo $a->slot instanceof ArrayObject ? 'mixed' : 'no';

__vybe_check(ob_get_clean(), "mixed");
