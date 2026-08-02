<?php
// vybe-test: php/generators_advanced2/generator_send_bidirectional
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

function accumulator(): Generator {
    $total = 0;
    while (true) {
        $n = yield $total;
        if ($n === null) break;
        $total += $n;
    }
}
$g = accumulator();
$g->current();
echo $g->send(5) . ',' . $g->send(3) . ',' . $g->send(10);

__vybe_check(ob_get_clean(), "5,8,18");
