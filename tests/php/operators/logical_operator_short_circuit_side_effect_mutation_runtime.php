<?php
// vybe-test: php/operators/logical_operator_short_circuit_side_effect_mutation_runtime
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

$hits = [];
$left = function() use (&$hits) {
    $hits[] = 'left';
    return false;
};
$right = function() use (&$hits) {
    $hits[] = 'right';
    return true;
};
echo false && $left();
echo '|';
echo 0 || $right();
echo '|';
echo count($hits);
echo '|';
echo implode(',', $hits);

__vybe_check(ob_get_clean(), "|1|1|right");
