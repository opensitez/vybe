<?php
// vybe-test: php/modern_php_deep/match_as_function_argument
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

function repeat(string $s, int $n): string { return str_repeat($s, $n); }
$x = 3;
echo repeat(match($x) { 1 => "a", 2 => "b", 3 => "c", default => "?" }, 4);

__vybe_check(ob_get_clean(), "cccc");
