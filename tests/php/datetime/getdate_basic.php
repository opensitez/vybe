<?php
// vybe-test: php/datetime/getdate_basic
// origin: languages/php/tests/php/test_datetime.rs

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

$ts = mktime(14, 30, 0, 6, 15, 2024);
$info = getdate($ts);
echo $info["year"];
echo $info["mon"];
echo $info["mday"];
echo $info["hours"];

__vybe_check(ob_get_clean(), "202461514");
