<?php
// vybe-test: php/php_session_regenerate_id_delete_old/test_php_session_regenerate_id_changes_session_id
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
$oldId = session_id();
@session_regenerate_id(false);
$newId = session_id();
@session_write_close();

echo $oldId !== $newId ? "REGENERATED_ID_OK" : "SAME_ID";

__vybe_check(ob_get_clean(), "REGENERATED_ID_OK");
