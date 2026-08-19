<?php
// vybe-test: php/literals/test_php_string_literals_and_escape_sequences
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

$s = 'hi\nthere';
echo $s;
echo '\n';
$name = 'B';
$t = "A{$name}";
echo $t;
echo '\n';
echo "abc\nxyz";

__vybe_check(ob_get_clean(), "hi\\nthere\\nAB\\nabc\nxyz");
