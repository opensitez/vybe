<?php
// vybe-test: php/loops/do_while_with_nested_if_and_stateful_continue
// origin: languages/php/tests/php/test_loops.rs

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

$i = 0;
$n = 0;
do {
    $i++;
    if ($i % 2 === 1) {
        continue;
    }
    $n++;
    if ($n === 2) { break; }
} while ($i < 10);
echo $i . '|' . $n;

__vybe_check(ob_get_clean(), "4|2");
