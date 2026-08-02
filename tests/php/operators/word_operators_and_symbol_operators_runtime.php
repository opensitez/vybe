<?php
// vybe-test: php/operators/word_operators_and_symbol_operators_runtime
// origin: languages/php/tests/php/test_operators.rs

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

$value = true;
if ($value and false or true) {
    echo 'word-and-or-1';
}
echo '|';
if (($value and false) or true) {
    echo 'word-and-or-2';
}
echo '|';
echo (true or false and false) . '|';
echo ((true or false) and false) . '|';
echo ($value && false or true) . '|';
echo ($value and false || true);

__vybe_check(ob_get_clean(), "word-and-or-1|word-and-or-2|1||1|1");
