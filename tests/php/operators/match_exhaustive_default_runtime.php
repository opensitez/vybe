<?php
// vybe-test: php/operators/match_exhaustive_default_runtime
// origin: languages/php/tests/php/test_operators.rs

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

$value = 3;
echo match ($value) {
    1 => 'one',
    2, 3 => 'two-or-three',
    default => 'other' };
echo '|';
echo match ($value > 1) {
    false => 'small',
    true => 'big' };
echo '|';
$list = [1, 2, 3];
echo match (true) {
    in_array($value, $list) => 'present',
    default => 'absent' };

__vybe_check(ob_get_clean(), "two-or-three|big|present");
