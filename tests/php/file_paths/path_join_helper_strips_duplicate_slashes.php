<?php
// vybe-test: php/file_paths/path_join_helper_strips_duplicate_slashes
// origin: languages/php/tests/php/test_file_paths.rs

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

function join_path(string $base, string $rel): string {
    return rtrim($base, '/') . '/' . ltrim($rel, '/');
}
echo join_path('/api/', 'users/1');

__vybe_check(ob_get_clean(), "/api/users/1");
