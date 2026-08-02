<?php
// vybe-test: php/php_exception_types_spl_hierarchy/test_php_spl_runtime_exception_subclasses
// origin: languages/php/tests/php/test_php_exception_types_spl_hierarchy.rs

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

function processBuffer(string $data) {
    if (strlen($data) === 0) throw new UnderflowException("Buffer empty");
    if (strlen($data) > 100) throw new OverflowException("Buffer full");
}

try {
    processBuffer("");
} catch (RuntimeException $e) { // UnderflowException extends RuntimeException
    echo "RuntimeException: " . get_class($e) . " -> " . $e->getMessage();
}

__vybe_check(ob_get_clean(), "RuntimeException: UnderflowException -> Buffer empty");
