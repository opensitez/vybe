<?php
// vybe-test: php/php_mbstring_multibyte_utf8_ops/test_php_mb_astral_literal_receiver
// origin: languages/php/tests/php/test_php_mbstring_multibyte_utf8_ops.rs

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

echo mb_strlen("a😀b€c") . " " . mb_substr("a😀b€c", 1, 1) . " " . mb_strpos("a😀b€c", "€") . " " . implode(",", mb_str_split("a😀b"));

__vybe_check(ob_get_clean(), "5 😀 3 a,😀,b");
