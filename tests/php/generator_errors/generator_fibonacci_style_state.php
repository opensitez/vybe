<?php
// vybe-test: php/generator_errors/generator_fibonacci_style_state
// origin: languages/php/tests/php/test_generator_errors.rs

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

function fib(int $n): Generator {
    $a = 0; $b = 1;
    for ($i = 0; $i < $n; $i++) {
        yield $a;
        [$a, $b] = [$b, $a + $b];
    }
}
echo implode(',', iterator_to_array(fib(4)));

__vybe_check(ob_get_clean(), "0,1,1,2");
