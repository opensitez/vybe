<?php
// vybe-test: php/php_enums_backed_tryfrom_cases/test_php81_backed_enum_tryfrom_returns_null_on_invalid
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

enum OrderState: string {
    case Pending = "pending";
    case Processing = "processing";
    case Completed = "completed";
}

$state1 = OrderState::tryFrom("completed");
$state2 = OrderState::tryFrom("invalid_state");

echo ($state1 !== null ? $state1->name : "NULL") . " | " . ($state2 === null ? "NULL" : "NOT_NULL");

__vybe_check(ob_get_clean(), "Completed | NULL");
