<?php
// vybe-test: php/host_mapped/curl_workflow
// origin: languages/php/tests/php/test_host_mapped.rs

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

echo "curl_workflow_ok";

__vybe_check(ob_get_clean(), "curl_workflow_ok");
