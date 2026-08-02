<?php
// vybe-test: php/sprintf_format_specifiers/sscanf_string_and_int
// origin: languages/php/tests/php/test_sprintf_format_specifiers.rs

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

[$name, $age] = sscanf("Alice 25", "%s %d");
echo "$name,$age";

__vybe_check(ob_get_clean(), "Alice,25");
