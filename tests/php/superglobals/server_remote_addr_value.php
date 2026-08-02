<?php
// vybe-test: php/superglobals/server_remote_addr_value
// origin: languages/php/tests/php/test_superglobals.rs

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

$_SERVER = ['REMOTE_ADDR' => '127.0.0.1'];
echo $_SERVER['REMOTE_ADDR'];

__vybe_check(ob_get_clean(), "127.0.0.1");
