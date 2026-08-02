<?php
// vybe-test: php/generator_send_runtime/generator_yield_object_by_reference
// origin: languages/php/tests/php/test_generator_send_runtime.rs

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

function g(): Generator {
    $o = new stdClass();
    $o->n = 1;
    yield $o;
}
$obj = iterator_to_array(g())[0];
echo $obj->n;

__vybe_check(ob_get_clean(), "1");
