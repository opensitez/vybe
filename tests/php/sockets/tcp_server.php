<?php
// vybe-test: php/sockets/tcp_server
// origin: languages/php/tests/php/test_sockets.rs

function __vybe_check($got, $want) {
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

echo "tcp_server_ok";

__vybe_check(ob_get_clean(), "tcp_server_ok");
