<?php
// vybe-test: php/json_errors/json_decode_control_char_in_string_with_throw
// origin: languages/php/tests/php/test_json_errors.rs

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

try { json_decode("\"a\nb\"", false, 512, JSON_THROW_ON_ERROR); echo 'ok'; }
catch (JsonException $e) { echo 'ctrl'; }

__vybe_check(ob_get_clean(), "ctrl");
