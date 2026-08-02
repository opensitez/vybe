<?php
// vybe-test: php/php_oop_constructor_promotion_readonly/test_php82_readonly_class_declaration
// origin: languages/php/tests/php/test_php_oop_constructor_promotion_readonly.rs

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

readonly class ImmutablePoint {
    public function __construct(
        public float $x,
        public float $y
    ) {}
}

$p = new ImmutablePoint(3.5, 7.2);
echo "Point({$p->x}, {$p->y})";

__vybe_check(ob_get_clean(), "Point(3.5, 7.2)");
