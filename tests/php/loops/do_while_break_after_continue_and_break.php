<?php
// vybe-test: php/loops/do_while_break_after_continue_and_break
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
$total = 0;
do {
    $i++;
    if ($i === 1) { continue; }
    $total += $i;
    if ($i === 4) { break; }
} while ($i < 6);
echo $total;

__vybe_check(ob_get_clean(), "9");
