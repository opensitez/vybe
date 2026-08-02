<?php
// vybe-test: php/callables/closure_bind_wrong_object_scope_silent_miss
// origin: languages/php/tests/php/test_callables.rs

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

class Alpha { private int $n = 1; }
class Beta {}
$fn = Closure::bind(function() { return $this->n ?? 0; }, new Beta(), Alpha::class);
echo $fn === null ? 'null' : (string)$fn();

__vybe_check(ob_get_clean(), "0");
