<?php
// vybe-test: php/generator_errors/generator_throw_then_resume_is_invalid
// origin: languages/php/tests/php/test_generator_errors.rs

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

function g(): Generator { yield 1; }
$gen = g();
$gen->next();
try {
    $gen->throw(new RuntimeException('boom'));
    $gen->next();
    echo 'resumed';
} catch (RuntimeException $e) {
    echo 'thrown';
}

__vybe_check(ob_get_clean(), "thrown");
