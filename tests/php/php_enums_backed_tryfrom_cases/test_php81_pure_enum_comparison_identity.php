<?php
// vybe-test: php/php_enums_backed_tryfrom_cases/test_php81_pure_enum_comparison_identity
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

enum Action { case Create; case Update; case Delete; }

$a1 = Action::Create;
$a2 = Action::Create;
$a3 = Action::Update;

echo ($a1 === $a2 ? "1" : "0");
echo ($a1 === $a3 ? "1" : "0");

__vybe_check(ob_get_clean(), "10");
