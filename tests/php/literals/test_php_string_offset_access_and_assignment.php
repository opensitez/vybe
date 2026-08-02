<?php
// vybe-test: php/literals/test_php_string_offset_access_and_assignment
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

$s = 'abc';
echo $s[0];
echo '\n';
echo $s[1];
$s[1] = 'Z';
echo '\n';
echo $s;

__vybe_check(ob_get_clean(), "a\nb\naZc");
