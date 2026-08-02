<?php
// vybe-test: php/php_enums_backed_methods_attributes/test_php81_string_backed_enum_from_tryfrom
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

enum Status: string {
    case Pending = "pending";
    case Active = "active";
}

$s = Status::from("active");
echo $s->name . "=" . $s->value;

__vybe_check(ob_get_clean(), "Active=active");
