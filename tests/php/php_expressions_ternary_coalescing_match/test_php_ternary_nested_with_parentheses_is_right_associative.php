<?php
// vybe-test: php/php_expressions_ternary_coalescing_match/test_php_ternary_nested_with_parentheses_is_right_associative
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

$a = 0;
echo ($a ? 'yes' : ($a ? 'inner-yes' : 'inner-no'));
echo '|';
echo ($a ? 'yes' : $a ? 'first' : 'second');

__vybe_check(ob_get_clean(), "inner-no|second");
