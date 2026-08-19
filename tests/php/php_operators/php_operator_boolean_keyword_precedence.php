<?php
// vybe-test: php/php_operators/php_operator_boolean_keyword_precedence
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

echo ((1 || 0 && 0)) ? 1 : 0, "\n";
echo (((1 || 0) && 0)) ? 1 : 0, "\n";
echo (function() { $a = true; $a = $a or false; return $a ? 1 : 0; })(), "\n";
echo (function() { $a = true; $a = $a || false; return $a ? 1 : 0; })(), "\n";
echo (function() { $a = false; $a = $a and true; return $a ? 1 : 0; })(), "\n";
echo (function() { $a = false; $a = $a && true; return $a ? 1 : 0; })(), "\n";


__vybe_check(ob_get_clean(), "1\n0\n1\n1\n0\n0");
