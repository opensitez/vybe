<?php
// vybe-test: php/try_catch_finally_return/finally_runs_on_foreach_continue
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
foreach ([1, 2, 3] as $n) {
    try {
        $log[] = "t$n";
        if ($n === 2) { continue; }
        $log[] = "a$n";
    } finally {
        $log[] = "f$n";
    }
}
echo implode(',', $log);

__vybe_check(ob_get_clean(), "t1,a1,f1,t2,f2,t3,a3,f3");
