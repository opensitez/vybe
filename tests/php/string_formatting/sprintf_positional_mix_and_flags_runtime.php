<?php
// vybe-test: php/string_formatting/sprintf_positional_mix_and_flags_runtime
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

echo sprintf('%2$s %1$s', 'left', 'right');
echo "\n";
echo sprintf('%1$+d', 42);
echo "\n";
echo sprintf('%1$ d', 42);
echo "\n";
echo sprintf('%1$010.2f', 3.5);

__vybe_check(ob_get_clean(), "right left\n+42\n42\n0000003.50");
