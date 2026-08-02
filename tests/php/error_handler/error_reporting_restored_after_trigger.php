<?php
// vybe-test: php/error_handler/error_reporting_restored_after_trigger
// origin: languages/php/tests/php/test_error_handler.rs

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

$before = error_reporting(E_ALL);
trigger_error('x', E_USER_NOTICE);
$after = error_reporting();
error_reporting($before);
echo $after === E_ALL ? 'all' : 'less';

__vybe_check(ob_get_clean(), "all");
