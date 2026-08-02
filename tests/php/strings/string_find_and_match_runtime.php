<?php
// vybe-test: php/strings/string_find_and_match_runtime
// origin: languages/php/tests/php/test_strings.rs

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

$s = 'hello world'; echo str_contains($s, 'world') ? '1' : '0'; echo "\n"; echo str_starts_with($s, 'he') ? '1' : '0'; echo "\n"; echo str_ends_with($s, 'ld') ? '1' : '0';

__vybe_check(ob_get_clean(), "1\n1\n1");
