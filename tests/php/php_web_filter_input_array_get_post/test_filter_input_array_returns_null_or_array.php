<?php
// vybe-test: php/php_web_filter_input_array_get_post/test_filter_input_array_returns_null_or_array
// origin: languages/php/tests/php/test_php_web_filter_input_array_get_post.rs

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

if (function_exists('filter_input_array')) {
    $res = filter_input_array(INPUT_GET, ['id' => FILTER_VALIDATE_INT]);
    echo ($res === null || is_array($res) || $res === false) ? 'filter_input_array_ok' : 'err', "\n";
} else {
    echo "filter_input_array_ok\n";
}

__vybe_check(ob_get_clean(), "filter_input_array_ok");
