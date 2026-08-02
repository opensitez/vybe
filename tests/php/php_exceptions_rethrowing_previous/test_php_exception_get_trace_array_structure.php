<?php
// vybe-test: php/php_exceptions_rethrowing_previous/test_php_exception_get_trace_array_structure
// origin: languages/php/tests/php/test_php_exceptions_rethrowing_previous.rs

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

function levelTwo() {
    throw new Exception("Deep Error");
}
function levelOne() {
    levelTwo();
}

try {
    levelOne();
} catch (Exception $e) {
    $trace = $e->getTrace();
    $funcs = array_column($trace, "function");
    echo "levelTwo=" . (in_array("levelTwo", $funcs) ? "1" : "0") . " levelOne=" . (in_array("levelOne", $funcs) ? "1" : "0");
}

__vybe_check(ob_get_clean(), "levelTwo=1 levelOne=1");
