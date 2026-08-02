<?php
// vybe-test: php/filter_var_domain_hostname/filter_var_domain_hostname_flag
// origin: languages/php/tests/php/test_filter_var_domain_hostname.rs

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

echo filter_var("http://example.com", FILTER_VALIDATE_DOMAIN, FILTER_FLAG_HOSTNAME) === false ? "fail" : "pass";

__vybe_check(ob_get_clean(), "fail");
