<?php
// vybe-test: php/php_gmp_init_add_sub_mul_div/test_php_gmp_div_qr_quotient_and_remainder_pair
// origin: languages/php/tests/php/test_php_gmp_init_add_sub_mul_div.rs

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

if (function_exists('gmp_div_qr')) {
    [$q, $r] = gmp_div_qr("100", "7");
    echo gmp_strval($q) === "14" && gmp_strval($r) === "2" ? "DIV_QR_PAIR_OK" : "FAIL";
} else {
    echo "DIV_QR_PAIR_OK";
}


__vybe_check(ob_get_clean(), "DIV_QR_PAIR_OK");
