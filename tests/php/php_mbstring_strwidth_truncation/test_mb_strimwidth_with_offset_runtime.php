<?php
// vybe-test: php/php_mbstring_strwidth_truncation/test_mb_strimwidth_with_offset_runtime
// origin: languages/php/tests/php/test_php_mbstring_strwidth_truncation.rs

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

echo mb_strimwidth('abcdef', 2, 4, '..'), "\n";
echo mb_strimwidth('日本語テスト', 1, 4, '..');

__vybe_check(ob_get_clean(), "cd..|本..");
