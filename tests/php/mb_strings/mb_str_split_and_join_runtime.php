<?php
// vybe-test: php/mb_strings/mb_str_split_and_join_runtime
// origin: languages/php/tests/php/test_mb_strings.rs

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

$s = "こんにちは";
$chunks = mb_str_split($s, 2);
echo count($chunks);
echo "|";
echo $chunks[0];
echo "|";
echo $chunks[1];

__vybe_check(ob_get_clean(), "3|こん|にち");
