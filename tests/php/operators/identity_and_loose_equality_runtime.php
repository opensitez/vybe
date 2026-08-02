<?php
// vybe-test: php/operators/identity_and_loose_equality_runtime
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

echo (1 == '1') ? 'eq1' : 'neq1';
echo (1 === '1') ? 'eq2' : 'neq2';
echo (0 == false) ? 'eq3' : 'neq3';
echo (0 === false) ? 'eq4' : 'neq4';
echo (null == false) ? 'eq5' : 'neq5';
echo (null === false) ? 'eq6' : 'neq6';

__vybe_check(ob_get_clean(), "eq1neq2eq3neq4eq5neq6");
