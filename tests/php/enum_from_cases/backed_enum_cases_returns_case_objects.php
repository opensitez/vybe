<?php
// vybe-test: php/enum_from_cases/backed_enum_cases_returns_case_objects
// origin: languages/php/tests/php/test_enum_from_cases.rs

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

enum Status: string { case Active = 'A'; case Inactive = 'I'; }
$names = array_map(fn($c) => $c->name, Status::cases());
sort($names);
echo implode(',', $names);

__vybe_check(ob_get_clean(), "Active,Inactive");
