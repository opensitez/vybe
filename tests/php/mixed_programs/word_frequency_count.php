<?php
// vybe-test: php/mixed_programs/word_frequency_count
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

$text = 'the quick brown fox jumps over the lazy dog the fox';
$words = explode(' ', $text);
$freq = array_count_values($words);
arsort($freq);
$top = array_slice($freq, 0, 2, true);
echo implode(',', array_map(fn($w,$c) => "$w:$c", array_keys($top), array_values($top)));

__vybe_check(ob_get_clean(), "the:3,fox:2");
