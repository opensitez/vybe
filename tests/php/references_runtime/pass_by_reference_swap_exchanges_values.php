<?php
// vybe-test: php/references_runtime/pass_by_reference_swap_exchanges_values
// origin: languages/php/tests/php/test_references_runtime.rs

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

function swap(string &$a, string &$b): void { [$a, $b] = [$b, $a]; }
$x = 'hi'; $y = 'bye';
swap($x, $y);
echo $x . ':' . $y;

__vybe_check(ob_get_clean(), "bye:hi");
