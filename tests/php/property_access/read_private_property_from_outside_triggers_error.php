<?php
// vybe-test: php/property_access/read_private_property_from_outside_triggers_error
// origin: languages/php/tests/php/test_property_access.rs

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

class Vault { private string $secret = 'hidden'; }
$v = new Vault();
try { echo $v->secret; }
catch (Error $e) { echo 'private'; }

__vybe_check(ob_get_clean(), "private");
