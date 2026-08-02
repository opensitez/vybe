<?php
// vybe-test: php/url_functions/rfc3986_path_join_pattern
// origin: languages/php/tests/php/test_url_functions.rs

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

function join_path(string $base, string $path): string {
    return rtrim($base, '/') . '/' . ltrim($path, '/');
}
echo join_path('https://example.com/api/', '/users');

__vybe_check(ob_get_clean(), "https://example.com/api/users");
