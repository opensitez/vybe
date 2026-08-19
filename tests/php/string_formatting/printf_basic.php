<?php
// vybe-test: php/string_formatting/printf_basic
// origin: languages/php/tests/php/test_string_formatting.rs

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

$written = printf("Name: %s, Age: %d\n", "Alice", 30);
echo $written > 0 ? 'wrote bytes' : 'nothing written';
echo "\n";


__vybe_check(ob_get_clean(), "Name: Alice, Age: 30\nwrote bytes");
