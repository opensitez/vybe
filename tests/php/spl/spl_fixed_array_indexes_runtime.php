<?php
// vybe-test: php/spl/spl_fixed_array_indexes_runtime
// origin: languages/php/tests/php/test_spl.rs

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

$fixed = new SplFixedArray(4);
$fixed[1] = 10;
$fixed[2] = 20;
echo $fixed->getSize();
echo '|';
echo $fixed->count();
echo '|';
echo $fixed[1];

__vybe_check(ob_get_clean(), "4|4|10");
