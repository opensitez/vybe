<?php
// vybe-test: php/vfprintf_stream_output/vfprintf_basic
// origin: languages/php/tests/php/test_vfprintf_stream_output.rs

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

$fp = fopen("php://memory", "w+");
$format = "Name: %s, Age: %d";
$args = ["Alice", 30];

$len = vfprintf($fp, $format, $args);
rewind($fp);
echo stream_get_contents($fp) . "|" . $len;
fclose($fp);

__vybe_check(ob_get_clean(), "Name: Alice, Age: 30|21");
