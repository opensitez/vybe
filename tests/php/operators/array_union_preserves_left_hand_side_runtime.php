<?php
// vybe-test: php/operators/array_union_preserves_left_hand_side_runtime
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

$left = ['first' => 'l', 2 => 'two'];
$right = [1 => 'one', 2 => 'override', 'extra' => 'x'];
$merged = $left + $right;
echo $merged['first'];
echo $merged[2];
echo $merged[1];
echo array_key_exists('extra', $merged) ? 'extra' : 'noextra';

__vybe_check(ob_get_clean(), "ltwooneextra");
