<?php
// vybe-test: php/union_types/param_type_class
// origin: languages/php/tests/php/test_union_types.rs

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

class Vec2 { public function __construct(public float $x, public float $y) {} }
function length(Vec2 $v): float { return sqrt($v->x**2 + $v->y**2); }
echo length(new Vec2(3.0, 4.0));

__vybe_check(ob_get_clean(), "5");
