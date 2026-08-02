<?php
// vybe-test: php/php_expressions_ternary_coalescing_match/test_php_match_with_computed_subject_and_boolean_arms
// origin: languages/php/tests/php/test_php_expressions_ternary_coalescing_match.rs

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

$left = 4;
$right = 2;
$result = match ($left > $right) {
    true => "gt",
    false => "le",
};
echo $result;
echo '|';
echo match (($left - $right) === 2) {
    true => "two",
    false => "not-two",
};

__vybe_check(ob_get_clean(), "gt|two");
