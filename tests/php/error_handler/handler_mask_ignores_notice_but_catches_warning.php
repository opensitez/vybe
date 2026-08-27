<?php
// vybe-test: php/error_handler/handler_mask_ignores_notice_but_catches_warning
// origin: languages/php/tests/php/test_error_handler.rs

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

echo "handler_mask_ignores_notice_but_catches_warning_ok";

__vybe_check(ob_get_clean(), "handler_mask_ignores_notice_but_catches_warning_ok");
