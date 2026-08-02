<?php
// vybe-test: php/try_catch_finally_return/finally_continue_do_while_loop
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
$n = 0;
do {
    try {
        $log[] = "t$n";
        if ($n === 0) { throw new Exception('skip'); }
    } catch (Exception $e) {
        $log[] = 'c';
    } finally {
        if ($n === 0) { $n++; continue; }
    }
    $n++;
} while ($n < 3);
echo implode(',', $log);

__vybe_check(ob_get_clean(), "t0,c,t1,t2");
