<?php
// vybe-test: php/mb_parse_str_population/mb_parse_str_basic
// origin: languages/php/tests/php/test_mb_parse_str_population.rs

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

mb_parse_str("first=value&arr[]=foo+bar&arr[]=baz", $output);
echo $output['first'] . "|" . $output['arr'][0] . "|" . $output['arr'][1];

__vybe_check(ob_get_clean(), "value|foo bar|baz");
