<?php
// vybe-test: php/php_array_is_list_validation/test_array_is_list_with_bool_zero_false_keys
// origin: languages/php/tests/php/test_php_array_is_list_validation.rs

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

$arr = [false => 'a', true => 'b'];
echo array_is_list($arr) ? 'true' : 'false', '|', count($arr);

__vybe_check(ob_get_clean(), "false|2");
