<?php
// vybe-test: php/strings/string_explode_preserves_empty_fields_runtime
// origin: languages/php/tests/php/test_strings.rs

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

$parts = explode(',', 'a,,b,');
echo count($parts);
echo '|';
echo $parts[1] === '' ? 'empty' : 'filled';
echo '|';
echo isset($parts[3]) ? ($parts[3] === '' ? 'tail-empty' : 'tail-fill') : 'tail-miss';

__vybe_check(ob_get_clean(), "4|empty|tail-empty");
