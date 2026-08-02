<?php
// vybe-test: php/php_expressions_match_nullsafe/test_php_nested_ternary_subject_and_match_default_only_one_branch
// origin: languages/php/tests/php/test_php_expressions_match_nullsafe.rs

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

$value = 2;
$result = match (true) {
    (($value > 0) ? 1 : 0) === 1 => 'positive',
    default => 'zero',
};
echo $result . '|';
echo match (($value > 0) ? $value : -$value) {
    2 => 'two',
    default => 'other',
};

__vybe_check(ob_get_clean(), "positive|two");
