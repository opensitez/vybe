<?php
// vybe-test: php/string_builtins_extended/str_word_count_modes_runtime
// origin: languages/php/tests/php/test_string_builtins_extended.rs

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

echo str_word_count("one two three");
echo "|";
echo implode(",", str_word_count("one two three", 1));
echo "|";
echo implode(",", array_keys(str_word_count("one two three", 2)));

__vybe_check(ob_get_clean(), "3|one,two,three|0,4,8");
