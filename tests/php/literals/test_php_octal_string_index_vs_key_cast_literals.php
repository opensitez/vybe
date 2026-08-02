<?php
// vybe-test: php/literals/test_php_octal_string_index_vs_key_cast_literals
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

$a = ['01' => 'string', 1 => 'one', 01 => 'zero'];
echo array_key_exists('1', $a) ? 'has1' : 'no1';
echo '|';
echo $a['1'];
echo '|';
echo $a[1];

__vybe_check(ob_get_clean(), "has1|zero|zero");
