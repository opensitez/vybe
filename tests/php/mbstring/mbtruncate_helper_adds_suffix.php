<?php
// vybe-test: php/mbstring/mbtruncate_helper_adds_suffix
// origin: languages/php/tests/php/test_mbstring.rs

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

function mb_truncate(string $s, int $max, string $suffix = '...'): string {
    if (mb_strlen($s) <= $max) return $s;
    return mb_substr($s, 0, $max - mb_strlen($suffix)) . $suffix;
}
echo mb_truncate('Hello World', 8);

__vybe_check(ob_get_clean(), "Hello...");
