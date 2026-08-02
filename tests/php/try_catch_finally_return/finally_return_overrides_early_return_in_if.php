<?php
// vybe-test: php/try_catch_finally_return/finally_return_overrides_early_return_in_if
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

function pick(bool $flag): string {
    try {
        if ($flag) { return 'early'; }
        return 'late';
    } finally {
        return 'always';
    }
}
echo pick(true) . pick(false);

__vybe_check(ob_get_clean(), "alwaysalways");
