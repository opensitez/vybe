<?php
// vybe-test: php/php_stream_socket_server_listen_accept/test_php_stream_get_transports_list
// origin: languages/php/tests/php/test_php_stream_socket_server_listen_accept.rs

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

$transports = stream_get_transports();
echo in_array("tcp", $transports) ? "TCP_AVAILABLE" : "NO_TCP";

__vybe_check(ob_get_clean(), "TCP_AVAILABLE");
