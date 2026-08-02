<?php
// vybe-test: php/oop_interfaces/interface_multiple_implementations
// origin: languages/php/tests/php/test_oop_interfaces.rs

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

interface Area { public function area(): float; }
class Rect implements Area { public function __construct(private float $w, private float $h) {} public function area(): float { return $this->w * $this->h; } }
class Circle implements Area { public function __construct(private float $r) {} public function area(): float { return round(M_PI * $this->r ** 2, 2); } }
echo (new Rect(3,4))->area() . ',' . (new Circle(1))->area();

__vybe_check(ob_get_clean(), "12,3.14");
