<?php
// vybe-test: php/numbers_extended/integer_cast_from_boolean_and_string
// origin: languages/php/tests/php/test_numbers_extended.rs

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

echo (int) true;
echo '|';
echo (int) false;
echo '|';
echo (int) "42";
echo '|';
echo (int) "04";

__vybe_check(ob_get_clean(), "1|0|42|4");
