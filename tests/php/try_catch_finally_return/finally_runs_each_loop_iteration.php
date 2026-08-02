<?php
// vybe-test: php/try_catch_finally_return/finally_runs_each_loop_iteration
// origin: languages/php/tests/php/test_try_catch_finally_return.rs

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

$log = [];
for ($i = 0; $i < 3; $i++) {
    try { $log[] = "b$i"; }
    finally { $log[] = "e$i"; }
}
echo implode(',', $log);

__vybe_check(ob_get_clean(), "b0,e0,b1,e1,b2,e2");
