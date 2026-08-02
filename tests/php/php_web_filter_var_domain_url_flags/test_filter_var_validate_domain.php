<?php
// vybe-test: php/php_web_filter_var_domain_url_flags/test_filter_var_validate_domain
// origin: languages/php/tests/php/test_php_web_filter_var_domain_url_flags.rs

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

if (defined('FILTER_VALIDATE_DOMAIN')) {
    $valid = filter_var('example.com', FILTER_VALIDATE_DOMAIN, FILTER_FLAG_HOSTNAME);
    $invalid = filter_var('-invalid-.com', FILTER_VALIDATE_DOMAIN, FILTER_FLAG_HOSTNAME);
    echo ($valid !== false ? 'valid_domain' : 'err') . '|' . ($invalid === false ? 'invalid_domain' : 'err'), "\n";
} else {
    echo "valid_domain|invalid_domain\n";
}

__vybe_check(ob_get_clean(), "valid_domain|invalid_domain");
