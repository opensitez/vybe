<?php
// vybe-test: php/exceptions/nested_try_catch_finally_runtime
// origin: languages/php/tests/php/test_exceptions.rs

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

function evalValue(int $v): int {
    try {
        if ($v < 0) {
            throw new Exception('negative');
        }
        return $v * 2;
    } catch (Exception $e) {
        return 0;
    } finally {
        return $v + 1;
    }
}
echo evalValue(-1) . '|' . evalValue(3);

__vybe_check(ob_get_clean(), "0|4");
