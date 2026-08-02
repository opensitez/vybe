<?php
// vybe-test: php/string_functions_extended/parse_str_boolean_strings_runtime
// origin: languages/php/tests/php/test_string_functions_extended.rs

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

parse_str('a=true&b=false&c=0&d=', $out);
echo $out['a'];
echo '|';
echo $out['b'];
echo '|';
echo $out['c'];
echo '|';
echo $out['d'];

__vybe_check(ob_get_clean(), "true|false|0|");
