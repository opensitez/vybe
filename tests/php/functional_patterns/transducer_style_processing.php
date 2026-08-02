<?php
// vybe-test: php/functional_patterns/transducer_style_processing
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

$data = range(1, 20);
$result = array_sum(
    array_slice(
        array_filter($data, fn($n) => $n % 3 === 0),
        0, 4
    )
);
echo $result;

__vybe_check(ob_get_clean(), "30");
