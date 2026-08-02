<?php
// vybe-test: php/match_edge_cases/nested_match_expressions
// origin: languages/php/tests/php/test_match_edge_cases.rs

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

$type = 'http';
$code = 200;
echo match($type) {
    'http' => match($code) { 200 => 'OK', 404 => 'Not Found', default => 'Other HTTP' },
    'ftp' => 'FTP',
    default => 'Unknown',
};

__vybe_check(ob_get_clean(), "OK");
