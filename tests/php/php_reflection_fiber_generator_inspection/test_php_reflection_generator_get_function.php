<?php
// vybe-test: php/php_reflection_fiber_generator_inspection/test_php_reflection_generator_get_function
// origin: languages/php/tests/php/test_php_reflection_fiber_generator_inspection.rs

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

function sampleGenerator() { yield "data"; }
$g = sampleGenerator();
$g->current();

$rg = new ReflectionGenerator($g);
$func = $rg->getFunction();
echo "FuncName: " . $func->getName();

__vybe_check(ob_get_clean(), "FuncName: sampleGenerator");
