<?php
// vybe-test: php/php_session_regenerate_id_delete_old/test_php_session_destroy_clears_session_file
// origin: languages/php/tests/php/test_php_session_regenerate_id_delete_old.rs

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
$_SESSION["auth"] = true;
$destroyed = @session_destroy();
echo $destroyed ? "DESTROYED_OK" : "FAIL";

__vybe_check(ob_get_clean(), "DESTROYED_OK");
