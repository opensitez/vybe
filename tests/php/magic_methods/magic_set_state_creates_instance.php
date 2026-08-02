<?php
// vybe-test: php/magic_methods/magic_set_state_creates_instance
// origin: languages/php/tests/php/test_magic_methods.rs

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
    public function __construct(public int $x, public int $y) {}
    public static function __set_state(array $props): self {
        return new self($props['x'], $props['y']);
    }
}
$p = new Point(3, 4);
echo $p->x;
echo $p->y;

__vybe_check(ob_get_clean(), "34");
