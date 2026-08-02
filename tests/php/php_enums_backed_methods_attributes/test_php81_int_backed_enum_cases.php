<?php
// vybe-test: php/php_enums_backed_methods_attributes/test_php81_int_backed_enum_cases
// origin: languages/php/tests/php/test_php_enums_backed_methods_attributes.rs

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

enum HTTPStatus: int {
    case OK = 200;
    case NotFound = 404;
}

$status = HTTPStatus::tryFrom(404);
echo $status ? $status->name : "NULL";

__vybe_check(ob_get_clean(), "NotFound");
