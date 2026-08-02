<?php
// vybe-test: php/variable_variables/variable_variables_with_nested_reference_names_runtime
// origin: languages/php/tests/php/test_variable_variables.rs

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

$field = 'state';
$prefix = 'app';
${$prefix . '_' . $field} = 'ready';
echo $app_state;
$bucket = [];
$bucket[$field] = 7;
$name = 'bucket';
echo $${$name}['state'];

__vybe_check(ob_get_clean(), "ready7");
