<?php
// vybe-test: php/generators_send_throw/generator_return_value_accessible_via_get_return
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

function sumToN(int $n): Generator {
    $sum = 0;
    for ($i = 1; $i <= $n; $i++) {
        $sum += $i;
        yield $i;
    }
    return $sum;
}
$gen = sumToN(4);
foreach ($gen as $_) {}
echo $gen->getReturn();

__vybe_check(ob_get_clean(), "10");
