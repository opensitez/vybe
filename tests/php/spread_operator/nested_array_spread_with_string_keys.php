<?php
// vybe-test: php/spread_operator/nested_array_spread_with_string_keys
// origin: languages/php/tests/php/test_spread_operator.rs

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

$left = ['base' => 'b'];
$mid = [...$left, 'mid' => 'm'];
$right = [...$mid, 'right' => 'r', 'base' => 'override'];
echo $right['base'];
echo '|';
echo $right['mid'];
echo '|';
echo $right['right'];

__vybe_check(ob_get_clean(), "override|m|r");
