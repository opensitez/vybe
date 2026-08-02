<?php
// vybe-test: php/php_web_http_response_headers_list_sent/test_headers_sent_file_line_vars
// origin: languages/php/tests/php/test_php_web_http_response_headers_list_sent.rs

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

if (function_exists('headers_sent')) {
    $file = '';
    $line = 0;
    $sent = headers_sent($file, $line);
    echo is_bool($sent) ? 'sent_bool_ok' : 'err', "\n";
} else {
    echo "sent_bool_ok\n";
}

__vybe_check(ob_get_clean(), "sent_bool_ok");
