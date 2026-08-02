<?php
// vybe-test: php/generators_advanced/lazy_range_generator_memory_efficient
// origin: languages/php/tests/php/test_generators_advanced.rs

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

function lazyRange(float $start, float $end, float $step = 1.0) {
    for ($i = $start; $i <= $end; $i += $step) {
        yield $i;
    }
}
$result = [];
foreach (lazyRange(0, 1, 0.25) as $v) {
    $result[] = $v;
}
echo implode(",", $result);

__vybe_check(ob_get_clean(), "0,0.25,0.5,0.75,1");
