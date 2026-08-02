<?php
// vybe-test: php/builtin_type_errors/array_keys_on_int_throws_type_error
// origin: languages/php/tests/php/test_builtin_type_errors.rs

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

try { array_keys(1); echo 'ok'; }
catch (TypeError $e) { echo 'keys-int'; }

__vybe_check(ob_get_clean(), "keys-int");
