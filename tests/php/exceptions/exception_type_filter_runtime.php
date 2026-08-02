<?php
// vybe-test: php/exceptions/exception_type_filter_runtime
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

try {
    throw new RuntimeException('runtime');
} catch (InvalidArgumentException $e) {
    echo 'arg';
} catch (RuntimeException $e) {
    echo $e->getMessage();
} catch (Exception $e) {
    echo 'base';
}

__vybe_check(ob_get_clean(), "runtime");
