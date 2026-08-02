<?php
// vybe-test: php/php_streams_file_system_wrapper_ops/test_php_pathinfo_filename_extension_directory
// origin: languages/php/tests/php/test_php_streams_file_system_wrapper_ops.rs

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

$path = "/var/www/html/index.controller.php";
$info = pathinfo($path);
echo $info["dirname"] . " | " . $info["basename"] . " | " . $info["extension"] . " | " . $info["filename"];

__vybe_check(ob_get_clean(), "/var/www/html | index.controller.php | php | index.controller");
