<?php
// vybe-test: php/modern_php_deep/readonly_class_deep
// origin: languages/php/tests/php/test_modern_php_deep.rs

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

readonly class Coordinate {
    public function __construct(
        public float $lat,
        public float $lng,
    ) {}
    public function distanceTo(Coordinate $other): float {
        return sqrt(($this->lat - $other->lat) ** 2 + ($this->lng - $other->lng) ** 2);
    }
}
$a = new Coordinate(0, 0);
$b = new Coordinate(3, 4);
echo $b->lat;
echo $a->distanceTo($b);

__vybe_check(ob_get_clean(), "35");
