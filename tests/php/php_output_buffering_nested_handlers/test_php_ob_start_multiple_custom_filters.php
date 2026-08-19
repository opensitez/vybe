<?php
// vybe-test: php/php_output_buffering_nested_handlers/test_php_ob_start_multiple_custom_filters
// origin: languages/php/tests/php/test_php_output_buffering_nested_handlers.rs

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

ob_start(fn($s) => str_replace("foo", "bar", $s));
ob_start(fn($s) => strtoupper($s));
echo "foo text";
ob_end_flush();
ob_end_flush();


__vybe_check(ob_get_clean(), "FOO TEXT");
