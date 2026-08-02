<?php
// vybe-test: php/sscanf_format_specifiers/sscanf_basic_return_array
// origin: languages/php/tests/php/test_sscanf_format_specifiers.rs

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

$str = "October 24, 1990";
$format = "%s %d, %d";
$res = sscanf($str, $format);
echo count($res) . "|" . $res[0] . "|" . $res[1] . "|" . $res[2];

__vybe_check(ob_get_clean(), "3|October|24|1990");
