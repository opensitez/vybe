<?php
// vybe-test: php/php_exceptions_rethrowing_previous/test_php_exception_get_file_line_code_getters
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

try {
    throw new Exception("Custom exception message", 404);
} catch (Exception $e) {
    echo "Code=" . $e->getCode() . " Msg=" . $e->getMessage() . " File=" . (strlen($e->getFile()) > 0 ? "OK" : "NO");
}

__vybe_check(ob_get_clean(), "Code=404 Msg=Custom exception message File=OK");
