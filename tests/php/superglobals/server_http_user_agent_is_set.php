<?php
// vybe-test: php/superglobals/server_http_user_agent_is_set
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

$_SERVER = ['HTTP_USER_AGENT' => 'TestBot/1.0'];
echo isset($_SERVER['HTTP_USER_AGENT']) ? 'ua' : 'none';

__vybe_check(ob_get_clean(), "ua");
