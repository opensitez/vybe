<?php
// vybe-test: php/php_stream_socket_client_timeout_options/test_php_stream_socket_client_timeout_setting
// origin: languages/php/tests/php/test_php_stream_socket_client_timeout_options.rs

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

$context = stream_context_create([
    "http" => ["timeout" => 2]
]);
$fp = @stream_socket_client("tcp://127.0.0.1:65534", $errno, $errstr, 1, STREAM_CLIENT_CONNECT, $context);
if ($fp) {
    stream_set_timeout($fp, 1);
    $info = stream_get_meta_data($fp);
    fclose($fp);
    echo "Timedout=" . ($info["timed_out"] ? "1" : "0");
} else {
    echo "Client connect failed (expected for closed port)";
}

__vybe_check(ob_get_clean(), "Client connect failed (expected for closed port)");
