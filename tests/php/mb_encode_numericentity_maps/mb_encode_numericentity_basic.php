<?php
// vybe-test: php/mb_encode_numericentity_maps/mb_encode_numericentity_basic
// origin: languages/php/tests/php/test_mb_encode_numericentity_maps.rs

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

$str = "—"; // EM DASH U+2014 (8212)
$convmap = [0x0, 0xffff, 0, 0xffff];
echo mb_encode_numericentity($str, $convmap, "UTF-8");

__vybe_check(ob_get_clean(), "&#8212;");
