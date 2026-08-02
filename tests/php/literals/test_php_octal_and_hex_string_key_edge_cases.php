<?php
// vybe-test: php/literals/test_php_octal_and_hex_string_key_edge_cases
// origin: languages/php/tests/php/test_literals.rs

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

$a = [0o10 => 'octal-key', 010 => 'legacy-octal', 0x10 => 'hex'];
echo $a[8];
echo '|';
echo $a[16];
echo '|';
echo $a['10'];

__vybe_check(ob_get_clean(), "octal-key|hex|octal-key");
