<?php
// vybe-test: php/operators/arithmetic_comparison_and_control_operator_runtime_results
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

echo 1 + 2;
echo 7 - 4;
echo 6 * 7;
echo 7 / 2;
echo 7 % 3;
echo 2 ** 3;
echo 'a' . 'b';
echo (-5) + 8;
echo (+5);
echo (!false) ? 't' : 'f';
echo (2 < 3) ? 't' : 'f';
echo (3 > 2) ? 't' : 'f';
echo (3 <= 3) ? 't' : 'f';
echo (4 >= 5) ? 't' : 'f';
echo (2 == '2') ? 't' : 'f';
echo (2 === '2') ? 't' : 'f';
echo (2 != 3) ? 't' : 'f';
echo (2 !== '2') ? 't' : 'f';
echo 1 <=> 2;
echo 2 <=> 2;
echo 3 <=> 2;
echo null ?? 'fallback';
echo 'value' ?? 'fallback';
echo false ? 'then' : 'else';
echo 0 ?: 'fallback';

__vybe_check(ob_get_clean(), "33423.518ab35ttttftftt-101fallbackvalueelsefallback");
