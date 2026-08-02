<?php
// vybe-test: php/generators_advanced/generator_variadic_throw_before_first_yield
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

function handled($head, ...$rest) {
    try {
        yield count($rest);
    } catch (Exception $e) {
        echo implode(',', $rest);
        yield $e->getMessage();
    }
}
$gen = handled('a', 'b', 'c');
echo $gen->throw(new Exception('stop'));

__vybe_check(ob_get_clean(), "b,cstop");
