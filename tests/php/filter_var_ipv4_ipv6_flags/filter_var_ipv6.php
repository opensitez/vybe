<?php
// vybe-test: php/filter_var_ipv4_ipv6_flags/filter_var_ipv6
// origin: languages/php/tests/php/test_filter_var_ipv4_ipv6_flags.rs

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

echo filter_var("2001:0db8:85a3:0000:0000:8a2e:0370:7334", FILTER_VALIDATE_IP, FILTER_FLAG_IPV6) ?: "fail";

__vybe_check(ob_get_clean(), "2001:0db8:85a3:0000:0000:8a2e:0370:7334");
