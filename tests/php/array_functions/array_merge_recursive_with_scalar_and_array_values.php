<?php
// vybe-test: php/array_functions/array_merge_recursive_with_scalar_and_array_values
// origin: languages/php/tests/php/test_array_functions.rs

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

$x = ['k' => 1];
$y = ['k' => [2, 3]];
$r = array_merge_recursive($x, $y);
echo is_array($r['k']) ? 'array' : 'scalar';
echo '|' . $r['k'][0];

__vybe_check(ob_get_clean(), "array|1");
