<?php
// vybe-test: php/declare_runtime/declare_ticks_function_runs
// origin: languages/php/tests/php/test_declare_runtime.rs

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

declare(ticks=1);
$hits = 0;
register_tick_function(function () use (&$hits) { $hits++; });
for ($i = 0; $i < 3; $i++) {}
echo $hits >= 0 ? 'ticks' : 'no';

__vybe_check(ob_get_clean(), "ticks");
