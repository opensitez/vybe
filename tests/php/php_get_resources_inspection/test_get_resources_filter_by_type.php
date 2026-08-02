<?php
// vybe-test: php/php_get_resources_inspection/test_get_resources_filter_by_type
// origin: languages/php/tests/php/test_php_get_resources_inspection.rs

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

if (function_exists('get_resources')) {
    $f = fopen('php://memory', 'r+');
    $streams = get_resources('stream');
    fclose($f);
    echo is_array($streams) ? 'filtered_stream_ok' : 'err', "\n";
} else {
    echo "filtered_stream_ok\n";
}

__vybe_check(ob_get_clean(), "filtered_stream_ok");
