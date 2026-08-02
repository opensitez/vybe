<?php
// vybe-test: php/operators/assignment_with_reference_like_chaining_runtime
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

$values = ['a' => ['b' => ['c' => 1]]];
$values['a']['b']['c'] += 4;
echo $values['a']['b']['c'];
echo '|';
$copy = $values;
$copy['a']['b']['c'] *= 2;
echo $copy['a']['b']['c'];
echo '|';
echo $values['a']['b']['c'];

__vybe_check(ob_get_clean(), "5|10|5");
