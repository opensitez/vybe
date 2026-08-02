<?php
// vybe-test: php/extra_100/interface_constant_in_class
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

interface Codes { const OK = 200; const NOT_FOUND = 404; }
class Response implements Codes {}
echo Response::OK . ',' . Response::NOT_FOUND;

__vybe_check(ob_get_clean(), "200,404");
