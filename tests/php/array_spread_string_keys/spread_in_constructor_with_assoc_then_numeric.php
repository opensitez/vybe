<?php
// vybe-test: php/array_spread_string_keys/spread_in_constructor_with_assoc_then_numeric
// origin: languages/php/tests/php/test_array_spread_string_keys.rs

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

class Pair {
    public function __construct(public int $a, public int $b, public int $c = 0) {}
}
$base = ['a' => 1];
$more = [2, 3];
$p = new Pair(...$base, ...$more);
echo $p->a . ',' . $p->b . ',' . $p->c;

__vybe_check(ob_get_clean(), "1,2,3");
