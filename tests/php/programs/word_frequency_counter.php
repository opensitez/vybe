<?php
// vybe-test: php/programs/word_frequency_counter
// origin: languages/php/tests/php/test_programs.rs

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

function wordFrequency(string $text): array {
    $words = str_word_count(strtolower($text), 1);
    $freq = [];
    foreach ($words as $w) $freq[$w] = ($freq[$w] ?? 0) + 1;
    arsort($freq);
    return $freq;
}
$freq = wordFrequency('the cat sat on the mat the cat');
echo $freq['the'] . "\n";
echo $freq['cat'] . "\n";
echo $freq['sat'] . "\n";

__vybe_check(ob_get_clean(), "3\n2\n1");
