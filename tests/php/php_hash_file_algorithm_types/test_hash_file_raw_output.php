<?php
// vybe-test: php/php_hash_file_algorithm_types/test_hash_file_raw_output
// origin: languages/php/tests/php/test_php_hash_file_algorithm_types.rs

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

$tmp = sys_get_temp_dir() . '/test_hash_file_raw.txt';
file_put_contents($tmp, 'data');
$raw = hash_file('md5', $tmp, true);
unlink($tmp);
echo strlen($raw), "\n";

__vybe_check(ob_get_clean(), "16");
