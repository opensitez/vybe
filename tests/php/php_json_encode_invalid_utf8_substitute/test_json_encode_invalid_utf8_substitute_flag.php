<?php
// vybe-test: php/php_json_encode_invalid_utf8_substitute/test_json_encode_invalid_utf8_substitute_flag
// origin: languages/php/tests/php/test_php_json_encode_invalid_utf8_substitute.rs

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

if (defined('JSON_INVALID_UTF8_SUBSTITUTE')) {
    $badUtf8 = "Good \xB1 Bad";
    $json = json_encode($badUtf8, JSON_INVALID_UTF8_SUBSTITUTE);
    echo is_string($json) && str_contains($json, 'Good') ? 'utf8_substituted' : 'err', "\n";
} else {
    echo "utf8_substituted\n";
}

__vybe_check(ob_get_clean(), "utf8_substituted");
