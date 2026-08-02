<?php
// vybe-test: php/php_oop_property_hooks_backing_field/test_property_hooks_backing_get_uses_virtual_projection
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

class Price {
    public int $cents = 1250 {
        get => $field / 100;
        set => (int) round($value * 100);
    }
}
$p = new Price();
$p->cents = 12.5;
echo $p->cents;

__vybe_check(ob_get_clean(), "12.5");
