<?php
// vybe-test: php/advanced_oop/polymorphic_method_dispatch
// origin: languages/php/tests/php/test_advanced_oop.rs

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

abstract class Shape {
    abstract public function area(): float;
    public function describe(): string { return get_class($this) . ':' . $this->area(); }
}
class Rect extends Shape { public function __construct(private float $w, private float $h) {} public function area(): float { return $this->w * $this->h; } }
class Triangle extends Shape { public function __construct(private float $b, private float $h) {} public function area(): float { return 0.5 * $this->b * $this->h; } }
$shapes = [new Rect(4,3), new Triangle(6,4)];
echo implode(',', array_map(fn($s) => $s->describe(), $shapes));

__vybe_check(ob_get_clean(), "Rect:12,Triangle:12");
