<?php
// vybe-test: php/named_arguments/named_args_in_match_arm_calls
// origin: languages/php/tests/php/test_named_arguments.rs

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

function clamp(int $val, int $min, int $max): int { return max($min, min($max, $val)); }
$values = [-5, 5, 15];
foreach ($values as $v) {
    $result = match(true) {
        $v < 0  => clamp(val: $v, min: 0, max: 10),
        $v > 10 => clamp(val: $v, min: 0, max: 10),
        default => $v };
    echo $result . "\n";
}

__vybe_check(ob_get_clean(), "0\n5\n10");
