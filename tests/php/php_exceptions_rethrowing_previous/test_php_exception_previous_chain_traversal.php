<?php
// vybe-test: php/php_exceptions_rethrowing_previous/test_php_exception_previous_chain_traversal
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

class DbError extends Exception {}
class ServiceError extends Exception {}

try {
    try {
        throw new DbError("Connection refused", 1001);
    } catch (DbError $e) {
        throw new ServiceError("Service unavailable", 503, $e);
    }
} catch (ServiceError $e) {
    $chain = [];
    $curr = $e;
    while ($curr !== null) {
        $chain[] = get_class($curr) . ": " . $curr->getMessage();
        $curr = $curr->getPrevious();
    }
    echo implode(" <- ", $chain);
}

__vybe_check(ob_get_clean(), "ServiceError: Service unavailable <- DbError: Connection refused");
