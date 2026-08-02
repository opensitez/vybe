<?php
// vybe-test: php/php_oop_property_hooks_asymmetric_visibility/test_php84_property_hooks_virtual_property
// origin: languages/php/tests/php/test_php_oop_property_hooks_asymmetric_visibility.rs

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

class Rectangle {
    public function __construct(public float $width, public float $height) {}

    public float $area {
        get => $this->width * $this->height;
    }
}

$r = new Rectangle(5.0, 4.0);
echo $r->area;

__vybe_check(ob_get_clean(), "20");
