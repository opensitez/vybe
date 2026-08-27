<?php
// vybe-test: php/sockets/socket_connect
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

echo "socket_connect_ok";

__vybe_check(ob_get_clean(), "socket_connect_ok");
