<?php
// vybe-test: php/mb_strings/mb_strlen_and_strlen_divergence_runtime
// origin: languages/php/tests/php/test_mb_strings.rs

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

$s = "café";
echo strlen($s);
echo "|";
echo mb_strlen($s);
echo "|";
echo mb_check_encoding($s, 'ASCII') ? "ascii-ok" : "ascii-no";

__vybe_check(ob_get_clean(), "5|4|ascii-no");
