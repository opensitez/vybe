<?php
// vybe-test: php/generators_advanced/generator_infinite_loop_broken_by_consumer
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

function incrementsForever(int $start = 0) {
    $n = $start;
    while (true) {
        yield $n++;
    }
}
$g = incrementsForever(1);
$result = [];
foreach ($g as $v) {
    $result[] = $v;
    if ($v >= 5) break;
}
echo implode(",", $result);

__vybe_check(ob_get_clean(), "1,2,3,4,5");
