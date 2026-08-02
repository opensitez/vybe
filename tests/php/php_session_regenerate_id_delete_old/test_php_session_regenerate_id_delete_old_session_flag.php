<?php
// vybe-test: php/php_session_regenerate_id_delete_old/test_php_session_regenerate_id_delete_old_session_flag
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
$_SESSION["user_id"] = 42;
$res = @session_regenerate_id(true); // true = delete old session data
echo "RegenerateResult=" . ($res ? "1" : "0") . " DataPreserved=" . ($_SESSION["user_id"] === 42 ? "YES" : "NO");
@session_write_close();

__vybe_check(ob_get_clean(), "RegenerateResult=1 DataPreserved=YES");
