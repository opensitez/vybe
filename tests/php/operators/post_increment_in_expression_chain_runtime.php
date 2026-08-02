<?php
// vybe-test: php/operators/post_increment_in_expression_chain_runtime
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

$counter = 1;
$left = $counter++;
$right = ++$counter;
echo $left . '|' . $right . '|' . $counter . '|' . ($left + $right);

__vybe_check(ob_get_clean(), "1|3|3|4");
