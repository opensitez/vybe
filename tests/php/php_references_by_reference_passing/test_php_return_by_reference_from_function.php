<?php
// vybe-test: php/php_references_by_reference_passing/test_php_return_by_reference_from_function
// origin: languages/php/tests/php/test_php_references_by_reference_passing.rs

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

$storage = 100;

function &getStorage(): int {
    global $storage;
    return $storage;
}

$ref = &getStorage();
$ref = 500;

echo $storage;

__vybe_check(ob_get_clean(), "500");
