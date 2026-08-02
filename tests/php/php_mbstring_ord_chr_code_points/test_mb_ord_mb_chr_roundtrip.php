<?php
// vybe-test: php/php_mbstring_ord_chr_code_points/test_mb_ord_mb_chr_roundtrip
// origin: languages/php/tests/php/test_php_mbstring_ord_chr_code_points.rs

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

$char = '𝄢'; // Musical symbol F clef
$code = mb_ord($char);
$restored = mb_chr($code);
echo ($char === $restored) ? 'roundtrip_ok' : 'err', "\n";

__vybe_check(ob_get_clean(), "roundtrip_ok");
