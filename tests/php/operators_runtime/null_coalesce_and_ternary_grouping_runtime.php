<?php
// vybe-test: php/operators_runtime/null_coalesce_and_ternary_grouping_runtime
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

$value = null;
echo ($value ?? 'null-fallback');
echo '|';
echo $value ? 'truthy' : ($value ?? 'or-default');
echo '|';
$value = '';
echo $value ?? 'none';
echo '|';
echo $value ?: 'fallback';

__vybe_check(ob_get_clean(), "null-fallback|or-default||fallback");
