<?php
// vybe-test: php/programs/anagram_detector
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

function isAnagram(string $a, string $b): bool {
    $sortStr = function(string $s): string {
        $chars = str_split(strtolower($s));
        sort($chars);
        return implode('', $chars);
    };
    return $sortStr($a) === $sortStr($b);
}
echo isAnagram('listen', 'silent') ? 'true' : 'false';
echo "\n";
echo isAnagram('hello', 'world') ? 'true' : 'false';
echo "\n";

__vybe_check(ob_get_clean(), "true\nfalse");
