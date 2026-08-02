<?php
// vybe-test: php/get_included_files_tracking/get_required_files_alias
// origin: languages/php/tests/php/test_get_included_files_tracking.rs

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

$files = get_required_files();
echo is_array($files) && count($files) >= 1 ? "ok" : "fail";

__vybe_check(ob_get_clean(), "ok");
