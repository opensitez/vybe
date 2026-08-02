<?php
// vybe-test: php/parse_str_array_population/parse_str_nested_arrays
// origin: languages/php/tests/php/test_parse_str_array_population.rs

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

$str = "user[name]=admin&user[roles][]=editor&user[roles][]=viewer";
parse_str($str, $output);
echo $output['user']['name'] . "|" . implode(',', $output['user']['roles']);

__vybe_check(ob_get_clean(), "admin|editor,viewer");
