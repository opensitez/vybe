<?php
// vybe-test: php/generators_advanced/yield_after_complex_expression
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

function transformedRange(int $n) {
    for ($i = 1; $i <= $n; $i++) {
        yield $i % 2 === 0
            ? $i * $i
            : $i * 2 + 1;
    }
}
$result = [];
foreach (transformedRange(6) as $v) {
    $result[] = $v;
}
echo implode(",", $result);
// i=1 odd: 1*2+1=3, i=2 even: 4, i=3 odd: 7, i=4 even: 16, i=5 odd: 11, i=6 even: 36

__vybe_check(ob_get_clean(), "3,4,7,16,11,36");
