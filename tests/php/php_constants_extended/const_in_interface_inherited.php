<?php
// vybe-test: php/php_constants_extended/const_in_interface_inherited
// origin: languages/php/tests/php/test_php_constants_extended.rs

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

interface HasMax { const int MAX = 100; }
interface HasMin extends HasMax { const int MIN = 0; }
class Range implements HasMin {}
echo Range::MAX . '-' . Range::MIN;

__vybe_check(ob_get_clean(), "100-0");
