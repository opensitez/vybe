<?php
// vybe-test: php/generators_patterns/generator_early_return
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

function limited(): Generator {
    yield 1; yield 2;
    if (true) return;
    yield 3;
}
echo implode(',', iterator_to_array(limited()));

__vybe_check(ob_get_clean(), "1,2");
