<?php
// vybe-test: php/enum_advanced_runtime/enum_cases_names_array
// origin: languages/php/tests/php/test_enum_advanced_runtime.rs

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

enum Color: string { case Red = 'r'; case Blue = 'b'; }
echo implode(',', array_map(fn($c) => $c->name, Color::cases()));

__vybe_check(ob_get_clean(), "Red,Blue");
