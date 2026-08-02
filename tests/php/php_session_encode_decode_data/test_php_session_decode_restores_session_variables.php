<?php
// vybe-test: php/php_session_encode_decode_data/test_php_session_decode_restores_session_variables
// origin: languages/php/tests/php/test_php_session_encode_decode_data.rs

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

@session_start();
$_SESSION = [];
$data = 'user|s:3:"Bob";id|i:42;';
@session_decode($data);

echo "User=" . ($_SESSION["user"] ?? "") . " ID=" . ($_SESSION["id"] ?? 0);
@session_write_close();

__vybe_check(ob_get_clean(), "User=Bob ID=42");
