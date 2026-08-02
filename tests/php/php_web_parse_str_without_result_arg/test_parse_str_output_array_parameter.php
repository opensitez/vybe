<?php
// vybe-test: php/php_web_parse_str_without_result_arg/test_parse_str_output_array_parameter
// origin: languages/php/tests/php/test_php_web_parse_str_without_result_arg.rs

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

$queryString = "a=10&b[]=x&b[]=y&c[name]=Alice";
parse_str($queryString, $result);
echo $result['a'] . '|' . implode(',', $result['b']) . '|' . $result['c']['name'], "\n";

__vybe_check(ob_get_clean(), "10|x,y|Alice");
