<?php
// vybe-test: php/magic_method_errors/dynamic_call_private_method_from_outside
// origin: languages/php/tests/php/test_magic_method_errors.rs

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

class Hidden { private function secret(): string { return 'no'; } }
$h = new Hidden();
try { $h->secret(); echo 'ok'; }
catch (Error $e) { echo 'hidden'; }

__vybe_check(ob_get_clean(), "hidden");
