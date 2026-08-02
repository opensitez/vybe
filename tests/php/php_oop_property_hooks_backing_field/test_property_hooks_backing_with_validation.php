<?php
// vybe-test: php/php_oop_property_hooks_backing_field/test_property_hooks_backing_with_validation
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

class Score {
    public int $value = 0 {
        set => max(0, $value);
    }
}
$s = new Score();
$s->value = -3;
echo $s->value;

__vybe_check(ob_get_clean(), "0");
