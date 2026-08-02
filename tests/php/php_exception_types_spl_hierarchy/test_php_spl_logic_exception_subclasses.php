<?php
// vybe-test: php/php_exception_types_spl_hierarchy/test_php_spl_logic_exception_subclasses
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

function checkAge(int $age) {
    if ($age < 0) throw new InvalidArgumentException("Age negative");
    if ($age > 150) throw new OutOfRangeException("Age out of bounds");
}

try {
    checkAge(-5);
} catch (LogicException $e) { // InvalidArgumentException extends LogicException
    echo "LogicException: " . get_class($e) . " -> " . $e->getMessage();
}

__vybe_check(ob_get_clean(), "LogicException: InvalidArgumentException -> Age negative");
