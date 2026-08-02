<?php
// vybe-test: php/operators/logical_and_or_xor_keyword_vs_symbol_runtime
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

$left = (true and false) ? 't' : 'f';
$right = (true && false) ? 't' : 'f';
echo $left . '|';
echo $right . '|';
echo (true xor false) ? 'x1' : 'x0';
echo '|';
echo ((bool) (0 && 1)) ? 'z1' : 'z0';

__vybe_check(ob_get_clean(), "f|f|x1|z0");
