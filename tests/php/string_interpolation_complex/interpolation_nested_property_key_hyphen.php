<?php
// vybe-test: php/string_interpolation_complex/interpolation_nested_property_key_hyphen
// origin: languages/php/tests/php/test_string_interpolation_complex.rs

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

$bag = ['user-name' => 'alice', 'meta' => ['tag-line' => 'dev']];
$payload = ['bag' => $bag];
echo "u: {$payload['bag']['user-name']}";
echo "|";
echo "m: {$payload['bag']['meta']['tag-line']}";

__vybe_check(ob_get_clean(), "u: alice|m: dev");
