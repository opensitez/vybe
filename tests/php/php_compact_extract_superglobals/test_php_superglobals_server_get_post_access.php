<?php
// vybe-test: php/php_compact_extract_superglobals/test_php_superglobals_server_get_post_access
// origin: languages/php/tests/php/test_php_compact_extract_superglobals.rs

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

$_SERVER["REQUEST_METHOD"] = "POST";
$_GET["page"] = "2";
$_POST["token"] = "abc123token";

echo $_SERVER["REQUEST_METHOD"] . " page=" . $_GET["page"] . " token=" . $_POST["token"];

__vybe_check(ob_get_clean(), "POST page=2 token=abc123token");
