<?php
// vybe-test: php/modern_php_deep/match_with_complex_arms
// origin: languages/php/tests/php/test_modern_php_deep.rs

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

function classify(int $score): string {
    return match(true) {
        $score >= 90 => "A",
        $score >= 80 => "B",
        $score >= 70 => "C",
        $score >= 60 => "D",
        default => "F" };
}
echo classify(95);
echo classify(82);
echo classify(55);

__vybe_check(ob_get_clean(), "ABF");
