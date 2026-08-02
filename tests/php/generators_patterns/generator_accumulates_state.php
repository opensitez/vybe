<?php
// vybe-test: php/generators_patterns/generator_accumulates_state
// origin: languages/php/tests/php/test_generators_patterns.rs

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

function runningTotal(array $nums): Generator {
    $total = 0;
    foreach ($nums as $n) { $total += $n; yield $total; }
}
echo implode(',', iterator_to_array(runningTotal([1,2,3,4,5])));

__vybe_check(ob_get_clean(), "1,3,6,10,15");
