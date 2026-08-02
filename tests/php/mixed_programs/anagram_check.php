<?php
// vybe-test: php/mixed_programs/anagram_check
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

function isAnagram(string $a, string $b): bool {
    $sort = function(string $s): string { $arr = str_split(strtolower($s)); sort($arr); return implode($arr); };
    return $sort($a) === $sort($b);
}
echo isAnagram('listen', 'silent') ? 'yes' : 'no';
echo ',';
echo isAnagram('hello', 'world') ? 'yes' : 'no';

__vybe_check(ob_get_clean(), "yes,no");
