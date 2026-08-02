<?php
// vybe-test: php/print_r_output/print_r_return_true_does_not_write_stdout
// origin: languages/php/tests/php/test_print_r_output.rs

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

$s = print_r('hidden', true);
echo 'marker:' . $s;

__vybe_check(ob_get_clean(), "marker:hidden");
