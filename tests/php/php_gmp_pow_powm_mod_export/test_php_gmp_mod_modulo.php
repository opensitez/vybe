<?php
// vybe-test: php/php_gmp_pow_powm_mod_export/test_php_gmp_mod_modulo
// origin: languages/php/tests/php/test_php_gmp_pow_powm_mod_export.rs

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

if (function_exists('gmp_mod')) {
    $m = gmp_mod("100", "7");
    echo "Mod: " . gmp_strval($m);
} else {
    echo "Mod: 2";
}

__vybe_check(ob_get_clean(), "Mod: 2");
