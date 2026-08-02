<?php
// vybe-test: php/programs/caesar_cipher_encode_decode
// origin: languages/php/tests/php/test_programs.rs

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

function caesarEncode(string $text, int $shift): string {
    $result = '';
    foreach (str_split($text) as $c) {
        if (ctype_upper($c)) $result .= chr((ord($c) - 65 + $shift) % 26 + 65);
        elseif (ctype_lower($c)) $result .= chr((ord($c) - 97 + $shift) % 26 + 97);
        else $result .= $c;
    }
    return $result;
}
$encoded = caesarEncode('Hello World', 3);
echo $encoded . "\n";
echo caesarEncode($encoded, 23) . "\n";

__vybe_check(ob_get_clean(), "Khoor Zruog\nHello World");
