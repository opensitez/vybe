<?php
// vybe-test: php/operators_runtime/logical_keyword_and_symbol_precedence_runtime
// origin: languages/php/tests/php/test_operators_runtime.rs

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

$x = false;
echo ($x and true || true) ? 't' : 'f';
echo '|';
echo ($x && true || true) ? 't' : 'f';
echo '|';
$y = false;
$y = false and true;
echo $y ? 't' : 'f';
echo '|';
$z = false && true;
echo $z ? 't' : 'f';

__vybe_check(ob_get_clean(), "f|t|f|f");
