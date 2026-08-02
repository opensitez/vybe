<?php
// vybe-test: php/variable_variables/variable_function_and_constant_from_namespace_runtime
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

namespace DynVarNS;

function make_value(string $value): string { return "value:$value"; }
const VALUE = 'v';
$factory = __NAMESPACE__ . '\\\\make_value';
echo $factory('x') . '|';
echo constant(__NAMESPACE__ . '\\\\VALUE');

__vybe_check(ob_get_clean(), "value:x|v");
