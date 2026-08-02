<?php
// vybe-test: php/php_oop_property_hooks_asymmetric_visibility/test_php84_private_set_blocked_from_outside_runtime
// origin: languages/php/tests/php/test_php_oop_property_hooks_asymmetric_visibility.rs

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

class Ledger {
    public private(set) int $balance = 0;
}
$l = new Ledger();
try {
    $l->balance = 10;
    echo 'wrote';
} catch (Error $e) {
    echo 'error';
}

__vybe_check(ob_get_clean(), "error");
