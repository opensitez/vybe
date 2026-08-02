<?php
// vybe-test: php/generator_get_return_value/generator_get_return_error_if_not_finished
// origin: languages/php/tests/php/test_generator_get_return_value.rs

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

function gen() {
    yield 1;
    return "done";
}
$g = gen();
try {
    echo $g->getReturn();
} catch (\Exception $e) {
    echo "error";
} catch (\Error $e) {
    echo "error";
}

__vybe_check(ob_get_clean(), "error");
