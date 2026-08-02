<?php
// vybe-test: php/patterns/prototype_clone_variation
// origin: languages/php/tests/php/test_patterns.rs

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

class Shape {
    public $color;
    public $x;
    public $y;
    public function __construct(string $color, int $x, int $y) {
        $this->color = $color;
        $this->x = $x;
        $this->y = $y;
    }
    public function move(int $dx, int $dy): self {
        $clone = clone $this;
        $clone->x += $dx;
        $clone->y += $dy;
        return $clone;
    }
}
$s1 = new Shape('red', 0, 0);
$s2 = $s1->move(5, 10);
echo $s1->x . ',' . $s1->y;
echo $s2->x . ',' . $s2->y;
echo $s2->color;

__vybe_check(ob_get_clean(), "0,05,10red");
