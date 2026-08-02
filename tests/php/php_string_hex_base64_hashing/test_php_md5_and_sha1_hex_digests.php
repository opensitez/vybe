<?php
// vybe-test: php/php_string_hex_base64_hashing/test_php_md5_and_sha1_hex_digests
// origin: languages/php/tests/php/test_php_string_hex_base64_hashing.rs

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

$str = "apple";
$md5 = md5($str);
$sha1 = sha1($str);

echo "md5_len=" . strlen($md5) . " sha1_len=" . strlen($sha1);

__vybe_check(ob_get_clean(), "md5_len=32 sha1_len=40");
