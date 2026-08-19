<?php
// vybe-test: php/string_formatting/printf_and_vprintf_return_value_runtime
// origin: languages/php/tests/php/test_string_formatting.rs

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

$a = printf("A=%d\n", 3);
echo "|";
$b = sprintf("%s %s", "x", "y");
echo "S=".$b;
echo "|";
$c = vprintf("%s:%d\n", ["p", 7]);
echo "|";
echo $a . "," . $c;

__vybe_check(ob_get_clean(), "A=3\n|S=x y|p:7\n|4,4");
