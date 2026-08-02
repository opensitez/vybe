<?php
// vybe-test: php/generators_send_throw/generator_return_value_is_null_if_no_return
// origin: languages/php/tests/php/test_generators_send_throw.rs

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

function simpleYield(): Generator {
    yield 1;
    yield 2;
}
$gen = simpleYield();
foreach ($gen as $_) {}
echo var_export($gen->getReturn(), true);

__vybe_check(ob_get_clean(), "NULL");
