<?php
// vybe-test: php/php_filter_input_get_post_cookie/test_php_filter_input_simulated_query_param
// origin: languages/php/tests/php/test_php_filter_input_get_post_cookie.rs

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

$_GET["age"] = "25";
$age = filter_input(INPUT_GET, "age", FILTER_VALIDATE_INT);
echo "Age: $age";

__vybe_check(ob_get_clean(), "Age: 25");
