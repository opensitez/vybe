<?php
// vybe-test: php/generators_send_throw/generator_lazy_range_with_step
// origin: languages/php/tests/php/test_generators_send_throw.rs

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

function lazyRange(int $start, int $end, int $step = 1): Generator {
    for ($i = $start; $i <= $end; $i += $step) {
        yield $i;
    }
}
echo implode(',', iterator_to_array(lazyRange(0, 10, 3)));

__vybe_check(ob_get_clean(), "0,3,6,9");
