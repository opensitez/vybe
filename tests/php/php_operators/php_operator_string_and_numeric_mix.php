<?php
// vybe-test: php/php_operators/php_operator_string_and_numeric_mix
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

echo ('a' . 1 . true), "\n";
echo ('a' . (1 + true)), "\n";
echo (1 + '2'), "\n";
echo ('3' + '4'), "\n";
echo ('2' . ('1' + 2)), "\n";
echo ('value=' . (1 ? 2 : 3)), "\n";
echo (function() { $left = 'left'; $right = null; return $left . ($right ?? 'fallback'); })(), "\n";
echo (function() { $left = 'left'; $right = 'right'; return $left . ($right ?? 'fallback'); })(), "\n";


__vybe_check(ob_get_clean(), "a11\na2\n3\n7\n23\nvalue=2\nleftfallback\nleftright");
