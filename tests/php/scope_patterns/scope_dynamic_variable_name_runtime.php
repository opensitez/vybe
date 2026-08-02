<?php
// vybe-test: php/scope_patterns/scope_dynamic_variable_name_runtime
// origin: languages/php/tests/php/test_scope_patterns.rs

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

$name = 'color';
$color = 'blue';
$$name = 'green';
echo $$name;
echo '|';
echo $color;

__vybe_check(ob_get_clean(), "green|blue");
