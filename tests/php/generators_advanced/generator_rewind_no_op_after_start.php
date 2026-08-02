<?php
// vybe-test: php/generators_advanced/generator_rewind_no_op_after_start
// origin: languages/php/tests/php/test_generators_advanced.rs

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

function once() {
    yield "first";
    yield "second";
}
$g = once();
echo $g->current();
$g->next();
echo $g->current();
// rewind on a started generator is a no-op / throws; we just verify
// we can call it without crashing by wrapping in try/catch
try {
    $g->rewind();
} catch (Exception $e) {
    echo "rewind-error";
}

__vybe_check(ob_get_clean(), "firstsecondrewind-error");
