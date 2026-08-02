<?php
// vybe-test: php/str_word_count_char_lists/str_word_count_char_lists
// origin: languages/php/tests/php/test_str_word_count_char_lists.rs

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

$str = "Hello fri3nd, you're looking good!";
echo count(str_word_count($str, 1, '3'));

__vybe_check(ob_get_clean(), "6");
