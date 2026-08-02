<?php
// vybe-test: php/php_reflection_parameter_is_passed_by_ref/test_reflection_parameter_is_passed_by_reference
// origin: languages/php/tests/php/test_php_reflection_parameter_is_passed_by_ref.rs

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

function modify_val(&$arg) { $arg = 100; }
$rf = new ReflectionFunction('modify_val');
$param = $rf->getParameters()[0];
echo $param->isPassedByReference() ? 'by_ref' : 'by_val', "\n";

__vybe_check(ob_get_clean(), "by_ref");
