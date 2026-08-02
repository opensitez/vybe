<?php
// vybe-test: php/oop_advanced/class_constant_expression
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

class Config {
    const BASE = 10;
    const DOUBLE = self::BASE * 2;
    const LABEL = "max:" . self::DOUBLE;
}
echo Config::BASE, "\n";
echo Config::DOUBLE, "\n";
echo Config::LABEL, "\n";

__vybe_check(ob_get_clean(), "10\n20\nmax:20");
