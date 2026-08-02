<?php
// vybe-test: php/literals/test_php_numeric_literal_mix_runtime
// origin: languages/php/tests/php/test_literals.rs

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

echo 0b10 + 0o10 + 0x10 + 1_000;
echo '|';
echo sprintf('%.1f', 1_000.5 + 2_000.25);
echo '|';
echo (int)'1_000';

__vybe_check(ob_get_clean(), "1034\n3000.8\n1");
