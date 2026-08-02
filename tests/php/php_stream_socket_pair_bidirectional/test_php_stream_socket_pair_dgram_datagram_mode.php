<?php
// vybe-test: php/php_stream_socket_pair_bidirectional/test_php_stream_socket_pair_dgram_datagram_mode
// origin: languages/php/tests/php/test_php_stream_socket_pair_bidirectional.rs

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

$pair = stream_socket_pair(STREAM_PF_UNIX, STREAM_SOCK_DGRAM, 0);
fwrite($pair[0], "Datagram Message");
$received = fread($pair[1], 1024);

fclose($pair[0]);
fclose($pair[1]);

echo $received;

__vybe_check(ob_get_clean(), "Datagram Message");
