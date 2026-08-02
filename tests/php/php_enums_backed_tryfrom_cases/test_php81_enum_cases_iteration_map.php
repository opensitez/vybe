<?php
// vybe-test: php/php_enums_backed_tryfrom_cases/test_php81_enum_cases_iteration_map
// origin: languages/php/tests/php/test_php_enums_backed_tryfrom_cases.rs

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

enum Level: int {
    case Low = 10;
    case Medium = 20;
    case High = 30;
}

$names = array_map(fn($case) => $case->name . "=" . $case->value, Level::cases());
echo implode(", ", $names);

__vybe_check(ob_get_clean(), "Low=10, Medium=20, High=30");
