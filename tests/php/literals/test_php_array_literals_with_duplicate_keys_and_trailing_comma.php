<?php
// vybe-test: php/literals/test_php_array_literals_with_duplicate_keys_and_trailing_comma
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

$values = ['a' => 1, 'b' => 2, 'a' => 3];
echo $values['a'];
echo '|';
echo $values['b'];

__vybe_check(ob_get_clean(), "3|2");
