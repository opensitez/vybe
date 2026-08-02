<?php
// vybe-test: php/named_args_extended/named_arg_skip_optional
// origin: languages/php/tests/php/test_named_args_extended.rs

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

function config(string $key, mixed $default = null, bool $required = false): mixed {
    return $required ? $key : ($default ?? $key);
}
echo config(key: 'host', required: true);

__vybe_check(ob_get_clean(), "host");
