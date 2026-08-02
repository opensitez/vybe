<?php
// vybe-test: php/generators_advanced/generator_return_value
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

function sumGenerator(array $numbers) {
    $sum = 0;
    foreach ($numbers as $n) {
        $sum += $n;
        yield $n;
    }
    return $sum;
}
$gen = sumGenerator([1, 2, 3, 4, 5]);
foreach ($gen as $val) {
    // consume all yielded values
}
echo $gen->getReturn();

__vybe_check(ob_get_clean(), "15");
