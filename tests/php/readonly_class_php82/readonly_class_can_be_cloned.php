<?php
// vybe-test: php/readonly_class_php82/readonly_class_can_be_cloned
// origin: languages/php/tests/php/test_readonly_class_php82.rs

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

readonly class Coord {
    public function __construct(public int $x, public int $y) {}
}
$a = new Coord(1, 2);
$b = clone $a;
echo $b->x . ',' . $b->y;

__vybe_check(ob_get_clean(), "1,2");
