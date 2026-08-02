<?php
// vybe-test: php/generator_current_key_sync/generator_current_key_sync
// origin: languages/php/tests/php/test_generator_current_key_sync.rs

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
    yield 'a' => 1;
    yield 'b' => 2;
}
$g = gen();
echo $g->key() . ":" . $g->current() . "|";
$g->next();
echo $g->key() . ":" . $g->current();

__vybe_check(ob_get_clean(), "a:1|b:2");
