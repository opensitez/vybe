<?php
// vybe-test: php/class_instantiation/private_constructor_blocks_external_new
// origin: languages/php/tests/php/test_class_instantiation.rs

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

class Vault {
    private function __construct() {}
    public static function open(): self { return new self(); }
}
try { new Vault(); echo 'ok'; }
catch (Error $e) { echo 'private'; }

__vybe_check(ob_get_clean(), "private");
