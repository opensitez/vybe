<?php
// vybe-test: php/generators/generator_send_injects_value_after_yield
// origin: languages/php/tests/php/test_generators.rs

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

function gen(): Generator {
    $x = yield 'first';
    yield 'got:' . $x;
}
$g = gen();
$g->current();
$g->send('Z');
echo $g->current();

__vybe_check(ob_get_clean(), "got:Z");
