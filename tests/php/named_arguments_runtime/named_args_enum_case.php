<?php
// vybe-test: php/named_arguments_runtime/named_args_enum_case
// origin: languages/php/tests/php/test_named_arguments_runtime.rs

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

enum E { case A; case B; }
function pick(E $e): string { return $e->name; }
echo pick(e: E::B);

__vybe_check(ob_get_clean(), "B");
