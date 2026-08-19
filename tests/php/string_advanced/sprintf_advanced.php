<?php
// vybe-test: php/string_advanced/sprintf_advanced
// origin: languages/php/tests/php/test_string_advanced.rs

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

echo sprintf("%05d", 42);
echo "\n";
echo sprintf("%.2f", 3.14159);
echo "\n";
echo sprintf("%s has %d items", "cart", 5);
echo "\n";
echo sprintf("%10s", "right");
echo "\n";
echo sprintf("%-10s|", "left");
echo "\n";

__vybe_check(ob_get_clean(), "00042\n3.14\ncart has 5 items\n     right\nleft      |");
