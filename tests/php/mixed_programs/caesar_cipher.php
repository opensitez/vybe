<?php
// vybe-test: php/mixed_programs/caesar_cipher
// origin: languages/php/tests/php/test_mixed_programs.rs

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

function caesar(string $text, int $shift): string {
    return preg_replace_callback('/[a-zA-Z]/', function($m) use ($shift) {
        $base = ctype_upper($m[0]) ? ord('A') : ord('a');
        return chr(($ord = ord($m[0]) - $base + $shift) % 26 >= 0 ? $base + $ord % 26 : $base + ($ord % 26 + 26));
    }, $text);
}
echo caesar('Hello World', 13);

__vybe_check(ob_get_clean(), "Uryyb Jbeyq");
