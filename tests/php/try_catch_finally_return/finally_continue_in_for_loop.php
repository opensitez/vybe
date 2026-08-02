<?php
// vybe-test: php/try_catch_finally_return/finally_continue_in_for_loop
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
for ($i = 0; $i < 4; $i++) {
    try {
        if ($i === 1) { throw new Exception('cont'); }
        $log[] = "ok$i";
    } finally {
        if ($i === 1) { continue; }
    }
}
echo implode(',', $log);

__vybe_check(ob_get_clean(), "ok0,ok2,ok3");
