<?php
// vybe-test: php/php_session_start_id_name/test_php_session_status_none_vs_active
// origin: languages/php/tests/php/test_php_session_start_id_name.rs

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

$before = session_status();
@session_start();
$after = session_status();
@session_write_close();

echo ($before === PHP_SESSION_NONE ? "NONE" : "OTHER") . " -> " . ($after === PHP_SESSION_ACTIVE ? "ACTIVE" : "OTHER");

__vybe_check(ob_get_clean(), "NONE -> ACTIVE");
