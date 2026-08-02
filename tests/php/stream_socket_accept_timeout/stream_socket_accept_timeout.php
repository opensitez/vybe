<?php
// vybe-test: php/stream_socket_accept_timeout/stream_socket_accept_timeout
// origin: languages/php/tests/php/test_stream_socket_accept_timeout.rs

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

$server = stream_socket_server("tcp://127.0.0.1:0", $errno, $errstr);
$start = microtime(true);
$conn = @stream_socket_accept($server, 0.1);
$end = microtime(true);

echo is_resource($server) ? "server|" : "no|";
echo is_resource($conn) ? "conn" : "timeout";

__vybe_check(ob_get_clean(), "server|timeout");
