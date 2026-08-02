<?php
// vybe-test: php/extra_100/generator_send_null_on_first_call
// origin: languages/php/tests/php/test_extra_100.rs

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

function gen():Generator{$v=yield 1;echo $v===null?'null':'notnull';}
$g=gen(); $g->current(); $g->next();

__vybe_check(ob_get_clean(), "null");
