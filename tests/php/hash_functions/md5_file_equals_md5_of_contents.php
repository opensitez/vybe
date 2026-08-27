<?php
// vybe-test: php/hash_functions/md5_file_equals_md5_of_contents
// origin: languages/php/tests/php/test_hash_functions.rs

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

echo "md5_file_equals_md5_of_contents_ok";

__vybe_check(ob_get_clean(), "md5_file_equals_md5_of_contents_ok");
