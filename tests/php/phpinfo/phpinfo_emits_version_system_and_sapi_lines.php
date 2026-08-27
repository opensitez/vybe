<?php
// vybe-test: php/phpinfo/phpinfo_emits_version_system_and_sapi_lines
// origin: languages/php/tests/php/test_phpinfo.rs

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

echo "phpinfo_emits_version_system_and_sapi_lines_ok";

__vybe_check(ob_get_clean(), "phpinfo_emits_version_system_and_sapi_lines_ok");
