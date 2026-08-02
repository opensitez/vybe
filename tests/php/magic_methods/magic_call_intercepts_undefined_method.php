<?php
// vybe-test: php/magic_methods/magic_call_intercepts_undefined_method
// origin: languages/php/tests/php/test_magic_methods.rs

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

class Proxy {
    public function __call($name, $args) {
        echo "called $name with " . count($args) . " args";
    }
}
$p = new Proxy();
$p->hello(1, 2, 3);

__vybe_check(ob_get_clean(), "called hello with 3 args");
