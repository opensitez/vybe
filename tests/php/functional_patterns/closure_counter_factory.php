<?php
// vybe-test: php/functional_patterns/closure_counter_factory
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

function makeCounter(int $start = 0): Closure {
    $n = $start;
    return fn() use (&$n) => ++$n;
}
$c1 = makeCounter();
$c2 = makeCounter(10);
echo $c1() . ',' . $c1() . ',' . $c2() . ',' . $c2();

__vybe_check(ob_get_clean(), "1,2,11,12");
