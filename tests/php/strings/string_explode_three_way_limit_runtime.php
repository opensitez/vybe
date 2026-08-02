<?php
// vybe-test: php/strings/string_explode_three_way_limit_runtime
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

$parts = explode('|', 'x|y|z|w', 3);
echo count($parts);
echo '|';
echo $parts[0];
echo '|';
echo $parts[1];
echo '|';
echo $parts[2];

__vybe_check(ob_get_clean(), "3|x|y|z|w");
