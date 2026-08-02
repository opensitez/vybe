<?php
// vybe-test: php/builtins/superglobal_string_key_read_after_assignment
// origin: languages/php/tests/php/test_builtins.rs

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

$_SERVER = ['SCRIPT_NAME' => '/cgi-bin/php.cgi']; echo $_SERVER['SCRIPT_NAME'];

__vybe_check(ob_get_clean(), "/cgi-bin/php.cgi");
