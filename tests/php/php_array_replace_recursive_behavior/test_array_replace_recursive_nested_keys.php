<?php
// vybe-test: php/php_array_replace_recursive_behavior/test_array_replace_recursive_nested_keys
// origin: languages/php/tests/php/test_php_array_replace_recursive_behavior.rs

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

$base = ['db' => ['host' => 'localhost', 'port' => 3306]];
$custom = ['db' => ['host' => '127.0.0.1', 'user' => 'root']];
$result = array_replace_recursive($base, $custom);
echo $result['db']['host'] . ':' . $result['db']['port'] . ':' . $result['db']['user'], "\n";

__vybe_check(ob_get_clean(), "127.0.0.1:3306:root");
