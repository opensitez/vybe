<?php
// vybe-test: php/functional_patterns/higher_order_sorter
// origin: languages/php/tests/php/test_functional_patterns.rs

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

function by(callable $key): Closure {
    return fn($a,$b) => $key($a) <=> $key($b);
}
$people = [['n'=>'Charlie','a'=>30],['n'=>'Alice','a'=>25],['n'=>'Bob','a'=>28]];
usort($people, by(fn($p) => $p['a']));
echo implode(',', array_column($people, 'n'));

__vybe_check(ob_get_clean(), "Alice,Bob,Charlie");
