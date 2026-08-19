<?php
// vybe-test: php/php_operators/php_operator_right_associative_examples
// origin: languages/php/tests/php/test_php_operators.rs

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

echo (1 <=> 2) > 0, "\n";
echo (2 <=> 1) > 0, "\n";
echo ((1 <=> 1) === 0), "\n";
echo ('1' <=> '2'), "\n";
echo ('a' <=> 'b'), "\n";
echo (true <=> false), "\n";
echo (false <=> true), "\n";


__vybe_check(ob_get_clean(), "\n1\n1\n-1\n-1\n1\n-1");
