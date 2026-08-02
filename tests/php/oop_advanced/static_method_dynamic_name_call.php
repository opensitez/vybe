<?php
// vybe-test: php/oop_advanced/static_method_dynamic_name_call
// origin: languages/php/tests/php/test_oop_advanced.rs

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

class Codec {
    public static function encode(string $v): string { return strtoupper($v); }
    public static function decode(string $v): string { return strtolower($v); }
}
$step = "encode";
$class = Codec::class;
echo $class::$step("ab");
echo "|";
$step = "decode";
echo $class::$step("XY");

__vybe_check(ob_get_clean(), "AB|xy");
