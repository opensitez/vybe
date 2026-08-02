<?php
// vybe-test: php/catch_type_union_order/catch_exception_when_error_subclass_thrown_stays_uncaught
// origin: languages/php/tests/php/test_catch_type_union_order.rs

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

$hit = '';
try { throw new AssertionError('fail'); }
catch (Exception $e) { $hit = 'exc'; }
catch (Error $e) { $hit = 'err'; }
echo $hit;

__vybe_check(ob_get_clean(), "err");
