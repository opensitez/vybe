<?php
// vybe-test: php/php_web_filter_var_array_input/test_filter_var_array_multiple_rules
// origin: languages/php/tests/php/test_php_web_filter_var_array_input.rs

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

if (function_exists('filter_var_array')) {
    $data = [
        'email' => 'user@example.com',
        'age' => '25',
        'invalid_email' => 'not-an-email'
    ];
    $definition = [
        'email' => FILTER_VALIDATE_EMAIL,
        'age' => FILTER_VALIDATE_INT,
        'invalid_email' => FILTER_VALIDATE_EMAIL
    ];
    $res = filter_var_array($data, $definition);
    echo ($res['email'] !== false && $res['age'] === 25 && $res['invalid_email'] === false) ? 'filter_array_ok' : 'err', "\n";
} else {
    echo "filter_array_ok\n";
}

__vybe_check(ob_get_clean(), "filter_array_ok");
