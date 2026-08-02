<?php
// vybe-test: php/php_web_session_id_custom_generator/test_session_create_id_prefix
// origin: languages/php/tests/php/test_php_web_session_id_custom_generator.rs

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

if (function_exists('session_create_id')) {
    $id = session_create_id('PREFIX-');
    echo str_starts_with($id, 'PREFIX-') ? 'prefix_ok' : 'err', "\n";
} else {
    echo "prefix_ok\n";
}

__vybe_check(ob_get_clean(), "prefix_ok");
