<?php
// vybe-test: php/modern_php_deep/readonly_property_in_constructor
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

class ImmutablePoint {
    public readonly float $x;
    public readonly float $y;
    public function __construct(float $x, float $y) {
        $this->x = $x;
        $this->y = $y;
    }
    public function distanceTo(ImmutablePoint $other): float {
        return sqrt(($this->x - $other->x) ** 2 + ($this->y - $other->y) ** 2);
    }
}
$a = new ImmutablePoint(0, 0);
$b = new ImmutablePoint(3, 4);
echo $b->x;
echo $a->distanceTo($b);

__vybe_check(ob_get_clean(), "35");
