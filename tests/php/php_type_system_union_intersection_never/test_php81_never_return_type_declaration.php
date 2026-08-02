<?php
// vybe-test: php/php_type_system_union_intersection_never/test_php81_never_return_type_declaration
// origin: languages/php/tests/php/test_php_type_system_union_intersection_never.rs

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

function stopExecution(string $msg): never {
    throw new RuntimeException($msg);
}

try {
    stopExecution("Halted");
} catch (RuntimeException $e) {
    echo "NEVER_RETURN: " . $e->getMessage();
}

__vybe_check(ob_get_clean(), "NEVER_RETURN: Halted");
