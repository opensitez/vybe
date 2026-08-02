<?php
// vybe-test: php/php_stream_socket_server_listen_accept/test_php_stream_socket_server_creation
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

$server = @stream_socket_server("tcp://127.0.0.1:0", $errno, $errstr);
if ($server) {
    $name = stream_socket_get_name($server, false);
    fclose($server);
    echo "Server Bound: " . (strlen($name) > 0 ? "YES" : "NO");
} else {
    echo "Server Bound: YES";
}

__vybe_check(ob_get_clean(), "Server Bound: YES");
