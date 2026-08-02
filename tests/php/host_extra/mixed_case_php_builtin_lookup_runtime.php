<?php
// vybe-test: php/host_extra/mixed_case_php_builtin_lookup_runtime
// origin: languages/php/tests/php/test_host_extra.rs

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

echo urlEncode('a b') === urlencode('a b') ? 'ok' : 'bad';
echo rawUrlEncode('c d') === rawurlencode('c d') ? 'ok' : 'bad';

__vybe_check(ob_get_clean(), "okok");
