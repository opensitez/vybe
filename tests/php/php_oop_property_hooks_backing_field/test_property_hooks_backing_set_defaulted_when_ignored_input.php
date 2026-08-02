<?php
// vybe-test: php/php_oop_property_hooks_backing_field/test_property_hooks_backing_set_defaulted_when_ignored_input
// origin: languages/php/tests/php/test_php_oop_property_hooks_backing_field.rs

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

class Counter {
    public int $count = 2 {
        set => $field + 1;
    }
}
$c = new Counter();
$c->count = 4;
$c->count = 0;
echo $c->count;

__vybe_check(ob_get_clean(), "3");
