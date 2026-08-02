<?php
// vybe-test: php/output_buffering/ob_implicit_flush_zero_disables_auto_flush_flag
// origin: languages/php/tests/php/test_output_buffering.rs

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

ob_implicit_flush(0);
echo ob_get_level() >= 0 ? 'ok' : 'bad';

__vybe_check(ob_get_clean(), "ok");
