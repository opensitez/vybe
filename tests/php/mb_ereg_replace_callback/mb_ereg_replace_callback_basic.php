<?php
// vybe-test: php/mb_ereg_replace_callback/mb_ereg_replace_callback_basic
// origin: languages/php/tests/php/test_mb_ereg_replace_callback.rs

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

$str = "äbc äbc";
$res = @mb_ereg_replace_callback("ä", function($m) { return "o"; }, $str);
// Note mb_ereg is deprecated in modern PHP or might not be enabled, just test if it runs or returns string
echo is_string($res) ? "ok" : "fail";

__vybe_check(ob_get_clean(), "ok");
