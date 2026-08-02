<?php
// vybe-test: php/generators_advanced2/fibonacci_generator
// origin: languages/php/tests/php/test_generators_advanced2.rs

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

function fibonacci(): Generator {
    [$a, $b] = [0, 1];
    while (true) { yield $a; [$a, $b] = [$b, $a + $b]; }
}
$gen = fibonacci();
$result = [];
for ($i = 0; $i < 8; $i++) { $result[] = $gen->current(); $gen->next(); }
echo implode(',', $result);

__vybe_check(ob_get_clean(), "0,1,1,2,3,5,8,13");
