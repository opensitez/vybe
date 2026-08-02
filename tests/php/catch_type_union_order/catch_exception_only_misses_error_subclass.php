<?php
// vybe-test: php/catch_type_union_order/catch_exception_only_misses_error_subclass
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

$log = [];
try {
    throw new DivisionByZeroError('div0');
} catch (Exception $e) {
    $log[] = 'exception';
} catch (Error $e) {
    $log[] = 'error';
}
echo implode(',', $log);

__vybe_check(ob_get_clean(), "error");
