<?php
// vybe-test: php/php_web_session_id_custom_generator/test_session_id_custom_setter_getter
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

$custom = 'custom_session_id_999';
$prev = session_id($custom);
echo session_id(), "\n";

__vybe_check(ob_get_clean(), "custom_session_id_999");
