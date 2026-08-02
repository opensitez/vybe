<?php
// vybe-test: php/type_juggling_runtime/string_floating_prefix_plus_suffix
// origin: languages/php/tests/php/test_type_juggling_runtime.rs

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

echo ("2.3" == 2.3) ? '1' : '0';
echo '|';
echo ("2.3" === 2.3) ? '1' : '0';
echo '|';
echo ("+5" == 5) ? '1' : '0';
echo '|';
echo ("5e2" == 500) ? '1' : '0';
echo '|';
echo ("5e2" === 500) ? '1' : '0';

__vybe_check(ob_get_clean(), "1|0|1|1|0");
