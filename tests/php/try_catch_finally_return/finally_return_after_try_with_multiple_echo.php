<?php
// vybe-test: php/try_catch_finally_return/finally_return_after_try_with_multiple_echo
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

function f(): string {
    try {
        echo 'a,';
        echo 'b,';
        return 'c';
    } finally {
        return 'd';
    }
}
echo f();

__vybe_check(ob_get_clean(), "a,b,d");
