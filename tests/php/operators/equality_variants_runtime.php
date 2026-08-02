<?php
// vybe-test: php/operators/equality_variants_runtime
// origin: languages/php/tests/php/test_operators.rs

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

echo (1 == '1') ? '1' : '0';
echo (1 === '1') ? '1' : '0';
echo (0 == false) ? '1' : '0';
echo (0 === false) ? '1' : '0';
echo ('0' == false) ? '1' : '0';
echo ('0' === false) ? '1' : '0';
echo ('' == false) ? '1' : '0';
echo ('' === false) ? '1' : '0';
echo ([] == false) ? '1' : '0';
echo ([] === false) ? '1' : '0';
echo ([] == null) ? '1' : '0';
echo ([1,2] == [1,2]) ? '1' : '0';
echo ([1,2] === [1,2]) ? '1' : '0';
echo ([1,2] === [2,1]) ? '1' : '0';
echo ('a' != 'b') ? '1' : '0';
echo ('a' <> 'a') ? '1' : '0';
echo (new stdClass() == new stdClass()) ? '1' : '0';
echo (new stdClass() === new stdClass()) ? '1' : '0';

__vybe_check(ob_get_clean(), "101010101011101010");
