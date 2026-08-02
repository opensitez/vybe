<?php
// vybe-test: php/programs/string_compression_run_length
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

function compress(string $s): string {
    if (strlen($s) === 0) return '';
    $result = '';
    $count = 1;
    for ($i = 1; $i <= strlen($s); $i++) {
        if ($i < strlen($s) && $s[$i] === $s[$i-1]) {
            $count++;
        } else {
            $result .= $s[$i-1] . $count;
            $count = 1;
        }
    }
    return $result;
}
echo compress('aabbc') . "\n";
echo compress('aaabbbccc') . "\n";
echo compress('abcd') . "\n";

__vybe_check(ob_get_clean(), "a2b2c1\na3b3c3\na1b1c1d1");
