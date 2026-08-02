<?php
// vybe-test: php/array_udiff_uassoc_callback/array_udiff_uassoc_empty_key_set
// origin: languages/php/tests/php/test_array_udiff_uassoc_callback.rs

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

$a = ["a" => "A", "b" => "B"];
$r = array_udiff_uassoc($a, ["A" => "A"], fn($v1,$v2) => strcmp($v1, $v2), "strcasecmp");
ksort($r);
echo count($r) . "|" . implode(',', array_keys($r));

__vybe_check(ob_get_clean(), "1|b");
