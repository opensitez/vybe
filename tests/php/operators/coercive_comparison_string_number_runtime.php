<?php
// vybe-test: php/operators/coercive_comparison_string_number_runtime
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

echo ("0" == 0) ? 'z' : 'nz';
echo '|';
echo ("0" === 0) ? 'zs' : 'nzs';
echo '|';
echo ("10" < 2) ? 'l' : 'g';
echo '|';
echo ("2" > 10) ? 'g2' : 'l2';

__vybe_check(ob_get_clean(), "z|nzs|g|l2");
