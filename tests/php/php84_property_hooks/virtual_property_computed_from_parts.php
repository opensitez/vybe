<?php
// vybe-test: php/php84_property_hooks/virtual_property_computed_from_parts
// origin: languages/php/tests/php/test_php84_property_hooks.rs

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

class Point {
    public function __construct(
        public float $x,
        public float $y,
    ) {}
    public float $distance {
        get => sqrt($this->x ** 2 + $this->y ** 2);
    }
}
$p = new Point(3, 4);
echo $p->distance;

__vybe_check(ob_get_clean(), "5");
