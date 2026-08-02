<?php
// vybe-test: php/php_reflection_generator_trace_inspection/test_reflection_generator_get_trace
// origin: languages/php/tests/php/test_php_reflection_generator_trace_inspection.rs

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

function worker() {
    yield 'step1';
}
$g = worker();
$g->current();
$rg = new ReflectionGenerator($g);
$trace = $rg->getTrace();
echo is_array($trace) ? 'trace_array_ok' : 'err', "\n";

__vybe_check(ob_get_clean(), "trace_array_ok");
