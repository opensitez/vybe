<?php
// vybe-test: php/try_catch_finally_return/inner_finally_runs_before_outer_finally_return
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
function f() use (&$log): string {
    try {
        try { $log[] = 'try'; return 'inner'; }
        finally { $log[] = 'inner_f'; return 'inner_ret'; }
    } finally {
        $log[] = 'outer_f';
        return 'outer_ret';
    }
}
echo f() . ':' . implode(',', $log);

__vybe_check(ob_get_clean(), "inner_ret:try,inner_f,outer_f");
