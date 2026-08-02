<?php
// vybe-test: php/operators_runtime/null_coalesce_skips_empty_string_runtime
// origin: languages/php/tests/php/test_operators_runtime.rs

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

$value = '';
echo ($value ?? 'fallback-1');
echo '|';
echo ($value ?: 'fallback-2');
echo '|';
$value = 0;
echo ($value ?? 'fallback-3');
echo '|';
echo ($value ?: 'fallback-4');

__vybe_check(ob_get_clean(), "|fallback-2|0|fallback-4");
