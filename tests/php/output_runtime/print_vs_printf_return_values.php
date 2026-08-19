<?php
// vybe-test: php/output_runtime/print_vs_printf_return_values
// origin: languages/php/tests/php/test_output_runtime.rs

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

echo print('p') . '|';
printf('%d', 7);
echo '|';
$n = printf('%s', 's');
echo $n;

__vybe_check(ob_get_clean(), "p|17|s1");
